# Sticky Notes for PowerPoint

Yellow sticky notes you can put right on top of your slides.

## Why

PowerPoint's built-in comments sit in a side pane, out of sight. On a busy deck it's easy to miss them, lose track of which slide they belong to, or forget to reply.

Real sticky notes work because you can't miss them. They sit right on the page, where the feedback applies. This add-in brings that to PowerPoint: a bright yellow note on the slide itself, easy to see, easy to answer, and easy to clear away when you're done.

## What it does

A **Sticky Notes** group on the Home tab:

| Button | What it does |
| --- | --- |
| **Add note** | Adds a yellow note to the top-right of the current slide. Click into it and type. |
| **Previous / Next** | Jumps to the previous or next slide that has a note, wrapping around at the ends. |
| **Delete** | Deletes the selected note, or all notes on the current slide if none is selected. |
| **Delete all** | Deletes every note in the deck, after asking for confirmation. |
| **Hide all / Show all** | Hides or shows every note in the deck, for example before presenting. |
| **Notes pane** | Opens a small side panel with the note count and messages. |

Notes are ordinary shapes saved in the deck. They travel with the file and are visible to anyone who opens it, even without the add-in. They also appear when you present or print, unless you hide them first.

## Requirements

- Microsoft 365 PowerPoint on **Windows, Mac or the web**
- Hide all / Show all need a recent build (PowerPoint JavaScript API 1.10)
- PowerPoint for iPad can show and edit notes, but the buttons aren't available there

## Privacy

The add-in runs entirely inside PowerPoint. Note text stays in your presentation file. Nothing is sent to this site or anywhere else; this repository only hosts the add-in's code and icons.

## Files

| File | Purpose |
| --- | --- |
| `taskpane.html` | The add-in page (side panel + ribbon commands) |
| `taskpane.css` | Panel styles |
| `notes.js` | All note logic |
| `assets/` | Ribbon icons |

## Installing

Sticky Notes is deployed to users by their Microsoft 365 administrator through **Admin center → Settings → Integrated apps**, using the add-in's manifest file. Once deployed, it appears on the Home tab in PowerPoint.

## Roadmap

- Choose where new notes appear (top right, top left, top center)
- Keyboard shortcuts
- A list of all notes in the pane, with click to jump
- Author, date and a "resolved" state for team review
