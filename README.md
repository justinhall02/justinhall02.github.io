# JHC Signature: Outlook add-in that inserts each staff member's email signature

Built 2026-10-24 (Justin: "Build it and I'll test it on my own"). Replaces the Exchange-rule and
CodeTwo/Exclaimer options (Rowan ask 416). Status: built and hosted, **not deployed to the tenant**.
Justin is testing it on his own mailbox first (ask "Test the Outlook signature add-in on your own mailbox").

## How it works

1. Outlook opens a new email, reply, reply all or forward. That fires the `OnNewMessageCompose` launch event.
2. Outlook starts the add-in in the background (no button, no pane). It runs `onNewMessageComposeHandler`
   in `docs/outlook-signature/commands.js`.
3. The handler reads `Office.context.mailbox.userProfile.emailAddress` and fetches
   `https://justinhall02.github.io/outlook-signature/signatures.json`.
4. It looks the email up (case-insensitive, `email` or any `aliases`). **Not in the file = it inserts nothing.**
   Fetch error, HTTP error or a 10 second timeout also = nothing. It always calls `event.completed()`.
5. If Outlook's own signature is on for that message, it turns it off (`disableClientSignatureAsync`) so
   there are not two, then inserts the JHC signature with `body.setSignatureAsync` (HTML). On a reply it goes
   above the quoted text.

The signature layout is the one built 2026-10-22 in
`personal-context/artifacts/2026-10/signatures-rollout/` (same text, colors, fonts, sizes and images). Cleanup
only: the Outlook-on-the-web clutter (`OWA` ids, `elementToProof`, the stray grey `FluentSystemIcons` line) is
gone, and the HubSpot click-tracking links were replaced by where they went (homepage, Justin's and Steve's
HubSpot meeting links, the Google review link, the Alarm.com dealer page). `TITLE_TBD` was not shipped: an empty
`title` leaves the title line out.

## Files

| Path | What |
|---|---|
| `docs/outlook-signature/manifest.xml` | The add-in manifest (XML "add-in only" type). This is what gets installed. |
| `docs/outlook-signature/commands.js` | The event handler + the signature template. One file, no imports. |
| `docs/outlook-signature/commands.html` | Loads Office.js + commands.js (web, new Outlook, Mac, mobile). |
| `docs/outlook-signature/signatures.json` | One entry per person. Public on purpose: names, titles, extensions only. |
| `docs/outlook-signature/assets/icon-*.png` | Add-in icons (16, 32, 64, 80, 128). |
| `docs/.well-known/microsoft-officeaddins-allowed.json` | Lets classic Outlook on Windows make the cross-site fetch (CORS). Must sit at the site root. |
| `docs/.nojekyll` | Stops GitHub Pages from dropping the `.well-known` folder. Do not delete. |
| `test/run-tests.js` | Mocked Office.js tests (Node vm with no DOM + headless Chrome). |
| `test/render-previews.js` | Renders each signature light and dark to PNG for design review. |

## Hosting

GitHub Pages, public, no login: repo **https://github.com/justinhall02/justinhall02.github.io** (local folder
`/mnt/c/Users/JustinHall/jhc-outlook-signature`), branch `main`, folder `/docs`, site
**https://justinhall02.github.io/**. The repo has to be the user site (named `justinhall02.github.io`)
because classic Outlook needs the `.well-known` file at the origin root, and a project site cannot serve the
root. Pages on the free plan needs a public repo; nothing secret lives here. Pages caches about 10 minutes, so a
JSON edit shows up in Outlook within ~10 minutes. Nothing runs on the VPS or ops.thejustinhall.com.

Manifest URL: https://justinhall02.github.io/outlook-signature/manifest.xml

## Which Outlook clients it covers

The XML add-in only manifest is used because it is the only manifest type that runs event-based add-ins on
every client. The unified (JSON) manifest does not run on Mac or iOS/Android.

| Client | Works | Minimum |
|---|---|---|
| Outlook on the web | Yes | none |
| New Outlook for Windows | Yes | none |
| Classic Outlook for Windows | Yes | Microsoft 365 build that supports Mailbox 1.10; Windows 10 1903+. Cross-site fetch needs Version 2201 (Build 16.0.14813.10000)+ |
| Outlook for Mac | Yes, **new Mac UI only** | Mailbox 1.10 (new Outlook for Mac) |
| Outlook for iOS / Android | Yes | Outlook 4.2352.0+; Microsoft 365 account |

Known gaps (from Microsoft's docs): opening an existing draft does not fire the event (by design: no double
insert); a reply started from the bottom of a message on mobile hides the signature until you go full screen;
iOS "Share" to mail does not fire it; "RSVP with note" and "forward a meeting" do not fire it; a shared mailbox or
send-as address still gets the signed-in user's signature (lookup is by `userProfile`, not `from`).

**Install rule that matters:** Microsoft says event-based add-ins launch on their own only when an admin
deploys them. A personal sideload works for testing (that is the documented test path), but for the whole team it
must be deployed from the Microsoft 365 admin center.

## Add a person (new hire)

1. Add one entry to `docs/outlook-signature/signatures.json`:
   ```json
   { "email": "first@justinhallconsulting.com", "aliases": [], "name": "First Last",
     "title": "Their Title", "ext": "105" }
   ```
   Optional: `"bookingUrl"` and `"reviewUrl"` (https only) add "Book time with me" / "Share your feedback on
   Google". `email` must equal what Outlook reports as the user's primary address.
2. Run the tests: `NODE_PATH=/home/justin/web-dev-agent/tools/node_modules node test/run-tests.js` (add the
   new email to `EXPECT` in the test).
3. Commit and push. Live in ~10 minutes. No manifest change, no admin step.

Change the design for everyone: edit `render()` in `commands.js`, re-run tests and `test/render-previews.js`,
run the design-reviewer agent on the PNGs, push.

## Deploy to everyone later (not done; needs Justin's go)

1. Microsoft 365 admin center > Settings > Integrated apps > Upload custom apps.
2. App type **Office Add-in**, "Provide link to manifest file":
   `https://justinhall02.github.io/outlook-signature/manifest.xml`.
3. Assign to specific users first (Justin, Steve, Jersey), then everyone. Takes up to 24 hours to appear;
   classic Outlook needs a restart.
4. Any later **manifest** change (not JSON or JS) needs the admin to accept the update. Bump `<Version>`.

## Rollback

- One person (sideload): Outlook on the web > Apps (or Get Add-ins) > My add-ins / Manage apps > JHC Signature > Remove.
  Then re-pick your own signature in Settings > Mail > Compose and reply (the add-in may have set it to none).
- Everyone (admin deploy): admin center > Integrated apps > JHC Signature > Remove app.
- Emergency, no admin access: empty the `people` list in `signatures.json` and push. The add-in then inserts
  nothing for anyone within ~10 minutes.

## Verify

- `npx office-addin-manifest validate docs/outlook-signature/manifest.xml`
- `NODE_PATH=/home/justin/web-dev-agent/tools/node_modules node test/run-tests.js` (`--live` uses the hosted JSON)
- `curl -sI` each hosted file logged-out: 200, right content type, `access-control-allow-origin: *`.
