# MNK client hospitality portal

Read-only booking history for Felipe, hosted by the Fika Google account.

## Security model

- Deployed as a separate Apps Script project, executing as the deploying Fika account.
- Felipe does not need a Google account. His private URL contains an access key.
- The access key is checked on the server and stored in Apps Script Properties, not in this repository.
- The portal deliberately shows all MNK bookings, regardless of who originally placed them.
- Only the fields explicitly returned by `sanitiseClientBooking_` reach the browser.
- This project contains no mutation, quote, calendar, cancellation or recharge functions.

## Deployment

1. Create a standalone Apps Script project and add its ID to `.clasp.json`.
2. Confirm the deploying Fika account can read the MNK dashboard spreadsheet.
3. In Apps Script Project Settings > Script Properties, create `FELIPE_PORTAL_KEY` with a long random value.
4. Add the deployed MNK booking-platform URL to `BOOKING_PLATFORM_URL` if the button is wanted.
5. Deploy as a web app: execute as the deploying Fika user; access set to anyone.
6. Give Felipe the deployment URL with `?key=THE_PRIVATE_VALUE` appended, or point a friendly redirect URL at it.
7. Test the private URL in a signed-out/incognito browser before sharing it.

If the link is forwarded, anyone holding it can view the portal. Rotate `FELIPE_PORTAL_KEY` to invalidate an old link.

After the one-time setup, run `felipepush.bat` from the repository root to push the files, update the existing deployment, commit the release and push the current Git branch.
