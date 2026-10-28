# Email Composer

Hosted files for the Email Composer Outlook add-in. One click in the Outlook compose window applies the house email format (Century Gothic 10pt, single spacing, one blank line around lists) to the text you have written. The quoted email below a reply is left unchanged.

This repository only serves the add-in's files through GitHub Pages. The source, tests and manifest are kept separately.

- `commands.html`, `commands.js`: the button's code
- `font-transform.js`: the formatting rules
- `assets/`: button icons

The add-in can only read and edit the email it is used on, makes no network requests, and cannot send email.
