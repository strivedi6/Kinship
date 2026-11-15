# Kinship: setup

`index.html` on GitHub Pages, data in the Google Sheet "Kinship", photos in the Drive folder "Kinship / Media".
Until step 2 is done the app runs in demo mode with a sample family.

Already created in your Drive folder "Kinship":
- Google Sheet "Kinship": https://docs.google.com/spreadsheets/d/1Bvwyl3cFtX0FYeyABic7NpvaivrX22V9Uabzq-GCEo4/edit
- Folder "Media": https://drive.google.com/drive/folders/1Nr_Q34Kr27e9RV5KPkL0VA6Ye0Ty8UnG

## 1. Apps Script (in the Kinship sheet)
1. Open the sheet > Extensions > Apps Script. Delete the sample code, paste all of `backend/Code.gs`, save.
2. Choose the function `setup` and press Run. Approve the permissions (Sheets, Drive, external requests for OneSignal, triggers).
   It builds the tabs, links the Media folder, writes your family code in the `Settings` tab and starts the 8 am birthday reminder.
3. Deploy > New deployment > Web app > Execute as **Me** > Who has access **Anyone** > Deploy. Copy the Web app URL (ends in `/exec`).

## 2. Connect the app
In `index.html`, near the top of the script, replace `PASTE_YOUR_APPS_SCRIPT_URL_HERE` with the Web app URL.

## 3. GitHub Pages
1. New repository, for example `kinship`.
2. The app files are in the root of this repository; `backend/Code.gs` is the Apps Script source (it does not run from GitHub).
3. Settings > Pages > Deploy from branch > `main` / root.

## 4. First sign-in
Open the Pages link, enter the family code, choose "I'm new here". **The first person to join becomes the admin**, so do this before sharing. Then: your photo > Share invite.

- iPhone: Safari > Share > Add to Home Screen.
- MacBook: Safari > File > Add to Dock (or Chrome > Install).
- Android: menu > Install app.

## 5. Notifications (OneSignal)
1. OneSignal > New App/Website > **Web** > choose **Custom Code**. Site URL = your Pages link (for example `https://strivedi6.github.io/kinship/`).
2. Copy the **App ID** and a **REST API key** (Settings > Keys & IDs).
3. In the sheet's `Settings` tab, paste them in the `value` column next to `oneSignalAppId` and `oneSignalKey`. The REST key stays in the sheet only and is never sent to the app.
4. Each person: open the app > their photo > Notifications > Turn on. On iPhone this only works when the app is opened from the Home Screen icon (iOS 16.4 or later).

What is sent: new post (to everyone), comment or love (to the post's author), new family date (to everyone), reward request (to the child's parents), reward decision (to the child), someone new joined (to admins), and at 8 am the day's birthdays and family dates.

## Chat, wishes and sign-in
- **Chat:** the chat button (top right, or the sidebar on a Mac) opens the Family group and one-to-one chats. A person only ever receives the group plus their own one-to-one messages. New messages arrive within about 5 seconds while Chat is open, and by notification otherwise.
- **Wishes:** Kids > Wishes. A child adds a wish, a parent sets how many stars it needs, and the child can ask for it once they have enough stars.
- **Sign-in:** family code once, then tap your name and enter your PIN on the keypad. New PINs are 4 digits.
- The `Messages` and `Wishes` tabs are created in the sheet automatically the first time they are used.

## Media folders
Every person gets their own subfolder in "Media" the first time they post a photo, named `Full Name (id)`. Photos stay private in Drive and are only sent to signed-in family. Profile photos are small thumbnails stored in the sheet.

## How it works
- **Sign-in:** family code + each person's own PIN (stored hashed). Five wrong PINs lock that profile for 15 minutes.
- **Tree:** each person sets father, mother and spouse in their profile. Grown-ups can add people who are not on the app; those people can claim their profile later by tapping their name at sign-in and choosing a PIN.
- **Kids:** parents add habits and rewards for their own children. Kids ask for a reward, a parent approves.
- **Roles:** admin, grown-up, kid (cannot add people, dates, habits or rewards).

## After editing Code.gs later
Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy. The URL stays the same.

## Design
iOS 27 UI kit: system font, large titles, floating glass tab bar on iPhone, sidebar on Mac, floating sheets, iOS system colours, light and dark mode. Motion and haptics follow Orbit: staggered screen entrance, spring sheets, a light tick on taps (vibration on Android, the system switch tick on iPhone).
