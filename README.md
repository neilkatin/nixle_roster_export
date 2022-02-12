
Nixle Export from google sheet rosters
======================================

This project takes our existing rosters (for CERT and ECC) and generates sheets that are compatible with nixle uploads,
to update the nixle alert database.

Once properly installed: there will be a new menu in google sheets ("Nixle").
If you select the "Generate Nixle Sheet" menu item the Nixle-<list> sheet will be updated (and created if it doesn't yet exist).

At that point you can download a CSV copy of the sheet (File -> Export -> Comma Separated Values) and upload it to Nixle.

Background info
---------------

Los Altos Hills CERT (Community Emergency Response Team) and ECC (Emergency Communications Committee) keep their rosters in google sheets.
We use [Nixle](https://www.nixle.com/) to send alerts to the team's cell phones.

This tool is to make it easy to update the Nixle data from the 'master' roster data.

Each team (CERT and ECC) have their own rosters in separate sheets.  In addition, CERT has two additional sub-rosters (Supervisors and Recon) that have their own Nixle groups.

This script takes care of generating all four Nixle inputs (CERT, ECC, Sups, Recon).

Quick dev instructions for this project
=======================================

[Instructions for using clasp to edit .gs source files locally](https://medium.com/geekculture/how-to-write-google-apps-script-code-locally-in-vs-code-and-deploy-it-with-clasp-9a4273e2d018)

I recommend installing without the -global option, and use npx to run clasp from the local repo.

* npx clasp login
* npx clasp clone *script-id*
* npx clasp pull
* npx clasp push -w
