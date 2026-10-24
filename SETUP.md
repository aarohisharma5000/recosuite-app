# Support form: setup (about 10 minutes, one time)

You'll set up a small free Google "Apps Script" that receives each ticket, saves it in a Google Sheet,
and emails you. No coding is needed. Just copy and paste.

## A. Create the ticket sheet and the script
1. Go to **sheets.google.com** (signed in as aarohisharma5000@gmail.com) and create a **new blank spreadsheet**.
   Name it **RecoSuite Support Tickets**.
   (Use a new sheet, not the subscriber sheet, so customer emails from tickets stay separate.)
2. In that sheet, click **Extensions → Apps Script**. A new tab opens with a file called `Code.gs`.
3. Delete everything in that file, then **paste the whole contents of `Code.gs`** from this folder.
4. Click the **💾 Save** icon. If asked for a project name, type **RecoSuite Support**.
   - Optional: to turn off the email you get for each ticket, change `var SEND_EMAIL_NOTIFICATION = true;` to `false` near the top, then save.

## B. Publish it as a Web App
5. Click the blue **Deploy** button (top right), then **New deployment**.
6. Click the ⚙️ gear next to "Select type" and choose **Web app**. Fill in:
   - Description: `Support form`
   - **Execute as: Me** (your Gmail)
   - **Who has access: Anyone** (not "Anyone with Google account". Customers must be able to send without logging in.)
7. Click **Deploy**, then **Authorize access**, choose your Google account, then **Advanced → Go to RecoSuite Support (unsafe) → Allow**.
   (Google shows this warning for every personal script. It's your own script. It asks to edit this spreadsheet and send email as you.)
8. **Copy the Web app URL.** It looks like `https://script.google.com/macros/s/AKfy…/exec`.
   - Check it: paste the URL into a new browser tab. You should see `{"ok":true,"service":"RecoSuite support",…}`.

## C. Connect the form
9. Open `support.html`. Near the bottom, find this line:
   `const APPS_SCRIPT_URL = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";`
   Replace the text between the quotes with your URL (keep the quotes), then save.
   (Or send the URL back and it can be filled in for you.)
10. Upload the new `support.html` to your GitHub Pages repo (the `recosuite` site), replacing the old one, and commit.
11. **Test:** open https://aarohisharma5000.github.io/recosuite/support.html, wait a few seconds, then send a test ticket.
    You should see "Ticket Submitted!". A new **Support** tab appears in your sheet with the ticket (time in IST, Status = Open),
    and you get an email. Replying to that email goes straight to the customer.

## Good to know
- **Changing the script later:** after editing `Code.gs`, go to **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**.
  The URL stays the same. If you skip this, the old version keeps running.
- **Don't** pick "New deployment" again unless you want a new URL. If you do, update `support.html`.
- **Status column:** change "Open" to "Done" (or anything else) as you work through tickets.
- If the script is ever unreachable, customers see a clear error with **WhatsApp** and **email** buttons that are already filled in,
  so no ticket is lost. The page never says "submitted" unless the ticket was actually saved.
- Spam protection: a hidden trap field plus a minimum time on the page, and at most 5 tickets per email every 10 minutes.
