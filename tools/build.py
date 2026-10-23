"""Write manifest.xml for the add-in, pointing at the given HTTPS base URL.

Usage: python3 tools/build.py https://leefc01.github.io/sticky-notes
The web files (taskpane.html, taskpane.css, notes.js, assets/) are served as-is from the repo root;
only manifest.xml is generated. Upload it in Microsoft 365 admin center > Integrated apps.
Bump VERSION whenever the manifest changes.
"""
import sys, shutil, pathlib, re
from xml.sax.saxutils import escape

ROOT = pathlib.Path(__file__).resolve().parent.parent
ADDIN_ID = "6f1c2b7e-4a3d-4c9e-9b52-5e0d8a7c31f4"   # keep stable across versions
VERSION = "1.2.0.0"                                   # = APP_VERSION in notes.js + ".0" (a test checks they match)
# 1.0.1.0: ribbon reduced to two controls, "Add note" button + "Notes" drop-down menu.
# 1.2.0.0 (v1.2): "Add Note" also in the Notes menu; title-case labels; free-slot placement.
# From v1.2 on, manifest version matches the product version (v1.0 = 1.0.0.0, v1.1 shipped as 1.0.1.0).

# Ribbon layout: PRIMARY is a standalone button; MENU_ITEMS sit in the "Notes" drop-down.
PRIMARY = ("Add", "addNote", "Add Note", "Add a sticky note to the current slide.", "add")
MENU = ("Menu", "Notes", "Jump between, hide, show or delete sticky notes.", "group")
MENU_ITEMS = [  # id, function (or None for pane), label, tooltip, icon
    ("MenuAdd", "addNote", "Add Note", "Add a sticky note to the current slide.", "add"),
    ("Prev", "previousNote", "Previous", "Go to the previous slide that has a note (wraps around).", "prev"),
    ("Next", "nextNote", "Next", "Go to the next slide that has a note (wraps around).", "next"),
    ("Delete", "deleteNote", "Delete", "Delete the selected note, or all notes on this slide if none is selected.", "delete"),
    ("DeleteAll", "deleteAllNotes", "Delete All", "Delete every note in this deck (asks first).", "deleteall"),
    ("Hide", "hideAllNotes", "Hide All", "Hide every note in this deck.", "hide"),
    ("Show", "showAllNotes", "Show All", "Show every note in this deck.", "show"),
    ("Pane", None, "Open Pane", "Open the Sticky Notes pane: note count, messages and confirmations.", "pane"),
]

def manifest(base):
    images, shorts, longs = [], [], []
    icons = sorted({b[4] for b in MENU_ITEMS} | {PRIMARY[4], MENU[3], "group"})
    for ic in icons:
        for px in (16, 32, 80):
            images.append(f'        <bt:Image id="Icon.{ic}.{px}" DefaultValue="{base}/assets/{ic}-{px}.png"/>')

    def icon(ic):
        return f'<Icon><bt:Image size="16" resid="Icon.{ic}.16"/><bt:Image size="32" resid="Icon.{ic}.32"/><bt:Image size="80" resid="Icon.{ic}.80"/></Icon>'

    def action(fn):
        return (f'<Action xsi:type="ExecuteFunction"><FunctionName>{fn}</FunctionName></Action>' if fn else
                '<Action xsi:type="ShowTaskpane"><SourceLocation resid="Taskpane.Url"/></Action>')

    def strings(bid, label, tip):
        shorts.append(f'        <bt:String id="{bid}.Label" DefaultValue="{escape(label)}"/>')
        longs.append(f'        <bt:String id="{bid}.Tip" DefaultValue="{escape(tip)}"/>')

    pid, pfn, plabel, ptip, pic = PRIMARY
    strings(pid, plabel, ptip)
    primary = f'''                <Control xsi:type="Button" id="SGNotes.{pid}">
                  <Label resid="{pid}.Label"/>
                  <Supertip><Title resid="{pid}.Label"/><Description resid="{pid}.Tip"/></Supertip>
                  {icon(pic)}
                  {action(pfn)}
                </Control>'''

    items = []
    for bid, fn, label, tip, ic in MENU_ITEMS:
        strings(bid, label, tip)
        items.append(f'''                    <Item id="SGNotes.{bid}">
                      <Label resid="{bid}.Label"/>
                      <Supertip><Title resid="{bid}.Label"/><Description resid="{bid}.Tip"/></Supertip>
                      {icon(ic)}
                      {action(fn)}
                    </Item>''')
    mid, mlabel, mtip, mic = MENU
    strings(mid, mlabel, mtip)
    nl = "\n"
    menu = f'''                <Control xsi:type="Menu" id="SGNotes.{mid}">
                  <Label resid="{mid}.Label"/>
                  <Supertip><Title resid="{mid}.Label"/><Description resid="{mid}.Tip"/></Supertip>
                  {icon(mic)}
                  <Items>
{nl.join(items)}
                  </Items>
                </Control>'''
    controls = [primary, menu]
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">
  <Id>{ADDIN_ID}</Id>
  <Version>{VERSION}</Version>
  <ProviderName>StatusGo</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sticky Notes"/>
  <Description DefaultValue="Yellow sticky notes on slides: add, jump between, hide, show and delete them."/>
  <IconUrl DefaultValue="{base}/assets/group-32.png"/>
  <HighResolutionIconUrl DefaultValue="{base}/assets/group-80.png"/>
  <SupportUrl DefaultValue="https://statusgo.ai"/>
  <Hosts>
    <Host Name="Presentation"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="{base}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <Runtimes>
          <Runtime resid="Taskpane.Url" lifetime="long"/>
        </Runtimes>
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="LearnMore.Url"/>
          </GetStarted>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="SGNotes.Group">
                <Label resid="Group.Label"/>
                <Icon><bt:Image size="16" resid="Icon.group.16"/><bt:Image size="32" resid="Icon.group.32"/><bt:Image size="80" resid="Icon.group.80"/></Icon>
{nl.join(controls)}
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
{nl.join(images)}
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="{base}/taskpane.html"/>
        <bt:Url id="LearnMore.Url" DefaultValue="https://statusgo.ai"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="Sticky Notes"/>
        <bt:String id="GetStarted.Title" DefaultValue="Sticky Notes is ready"/>
{nl.join(shorts)}
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Use the Sticky Notes group on the Home tab to add and manage notes."/>
{nl.join(longs)}
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
'''

def main():
    if len(sys.argv) != 2 or not re.match(r"^https://[^\s]+[^/]$", sys.argv[1]):
        sys.exit("Usage: build.py https://host/path   (HTTPS, no trailing slash)")
    base = sys.argv[1]
    (ROOT / "manifest.xml").write_text(manifest(base), encoding="utf-8")
    print(f"Wrote manifest.xml v{VERSION} for {base}")

if __name__ == "__main__":
    main()
