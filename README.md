# Deal Team Scheduler

An Outlook add-in that saves named groups of colleagues ("deal teams") and adds a whole team to a meeting invite in one click. After adding a team, open the Scheduling Assistant to see everyone's availability.

Your teams are stored in Outlook roaming settings, which live inside your own mailbox. The add-in has no backend, makes no calls to Microsoft Graph or any other service, and needs no app registration.

## What's in this folder

- `docs/` holds the add-in itself (the task pane page, styles, script, and icons). This is the folder you host.
- `manifest.template.xml` is the add-in manifest with a placeholder for your hosting URL.
- `scripts/build-manifest.js` fills in your URL and writes `manifest.xml`, the file you upload to Outlook.

There is no build step. The files in `docs/` are served exactly as they are.

## Step 1: Check that your Outlook allows custom add-ins

In Outlook on the web, go to https://aka.ms/olksideload and open My add-ins. If you see a Custom add-ins section with an option to add a custom add-in, you can continue. If not, your IT admin has turned this off and you will need their approval first.

## Step 2: Host the `docs` folder on GitHub Pages

Outlook loads the add-in from an HTTPS web address each time you open it. GitHub Pages is free and works well. Only the add-in code is hosted there; your deal teams never leave your mailbox.

1. Sign in at github.com (create a free account if you need one).
2. Create a new repository named `deal-team-scheduler`. Set it to Public (Pages on private repositories requires a paid plan).
3. On the new repository page, choose "uploading an existing file." Drag in everything from this project folder, including the `docs` folder, and commit.
4. In the repository, go to Settings, then Pages. Under "Build and deployment," set Source to "Deploy from a branch," choose the `main` branch and the `/docs` folder, and save.
5. After a minute or two, your add-in will be live at `https://YOUR-GITHUB-USERNAME.github.io/deal-team-scheduler`. Open `https://YOUR-GITHUB-USERNAME.github.io/deal-team-scheduler/taskpane.html` in a browser to confirm it loads. It will say it needs to run inside Outlook, which is expected.

## Step 3: Create your manifest

In Terminal, from this project folder:

```
npm run manifest -- https://YOUR-GITHUB-USERNAME.github.io/deal-team-scheduler
npm run validate
```

The first command writes `manifest.xml`. The second checks it against Microsoft's validator and should report that the manifest is valid.

## Step 4: Upload the manifest to Outlook

Outlook on the web and new Outlook share the same add-in list, so installing once covers both.

1. Go to https://aka.ms/olksideload.
2. Open My add-ins, then under Custom add-ins choose Add a custom add-in, then Add from file.
3. Select `manifest.xml` and confirm the install prompt.

In classic Outlook for Mac or Windows, the same option is under Get Add-ins, then My add-ins, then Add a custom add-in, then Add from file. Add-ins installed on the web usually appear in classic Outlook automatically after a restart.

## Using it

1. Create a new meeting (or open one you organize).
2. Click Deal Teams in the ribbon. In new Outlook and on the web, it may be under the Apps button.
3. Choose New team, name it after the deal, paste in email addresses, and save. You can paste straight from an Outlook To line, including entries like `Doe, Jane <jane.doe@company.com>`.
4. Open the team, uncheck anyone you want to leave off this time, and click Add as required or Add as optional.
5. Open the Scheduling Assistant to see everyone's availability.

Use "Back up or move your teams" at the bottom of the team list to export your teams as text. Keep a copy before removing or reinstalling the add-in.

## Making changes later

Edit the files in `docs/`, then upload the changed files to the same GitHub repository. Outlook picks up the new version the next time the pane opens, with no need to re-upload the manifest. You only need a new manifest if the hosting URL changes; in that case, run Step 3 again, remove the old add-in in Outlook, and upload the new manifest.

## Testing locally (optional)

If you want to try changes before publishing:

```
npm run certs
npm run manifest:local
npm start
```

Then upload the generated `manifest.xml` as in Step 4. The add-in will only work while `npm start` is running on your Mac. Switch back with Step 3 when you are done.

## Limits

- Outlook caps add-in settings at about 32 KB, which is roughly a few hundred people across all teams. The add-in warns you when you get close.
- The Deal Teams button appears on meetings you are organizing. It does not appear on meetings you were invited to, since you can't change those attendee lists.
