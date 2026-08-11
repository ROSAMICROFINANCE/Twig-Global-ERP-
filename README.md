# HardwareERP

A point-of-sale + inventory system: Quotations, Sales Orders, Purchase Orders,
Stock, Goods Returns, Expenses, Reports, and Accounting — with per-user logins
and module-level rights.

## How passwords are hardened

Passwords are no longer stored in this app at all. Signing in and creating
accounts is handled by **Firebase Authentication**, which hashes and salts
passwords on Google's servers — the same system used by real production
apps. This app only ever sees a signed-in session, never a raw password
after account creation.

## 1. Create your Firebase project (free tier is enough for 5 users)

1. Go to https://console.firebase.google.com and create a new project.
2. In the project, click **Build → Authentication → Get started**, then
   enable the **Email/Password** sign-in method.
3. Click **Build → Firestore Database → Create database**, start in
   **production mode**, pick a region close to you.
4. In **Project settings → General**, scroll to "Your apps", click the
   `</>` (web) icon, register an app (any nickname), and copy the
   `firebaseConfig` object it gives you.
5. Paste those values into `src/firebase.js` in this project, replacing the
   placeholder strings.
6. In Firestore, open the **Rules** tab and paste in the contents of
   `firestore.rules` from this project, then Publish.

## 2. Run it locally to test

```bash
npm install
npm run dev
```

Open the local URL it prints. The first screen will ask you to create the
admin account — that's you. After that, use **Settings** (visible only to
the admin) to add your 5 staff logins and tick which modules each one can
see.

## 3. Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

(Create the empty repo on GitHub first, then run the commands above from
this project folder.)

## 4. Put it on the web — two good options

**Option A: Firebase Hosting (simplest, same project as your data)**

```bash
npm install -g firebase-tools
firebase login
firebase init hosting
# When asked for your public directory, enter: dist
# When asked "configure as a single-page app", answer: Yes
npm run build
firebase deploy
```

It will print a live `https://your-project.web.app` URL — that's your
shop's app, reachable from any phone or computer.

**Option B: Connect GitHub to Vercel or Netlify (auto-deploys on every push)**

1. Go to vercel.com or netlify.com, sign in with GitHub, and import this
   repository.
2. Build command: `npm run build`  ·  Output directory: `dist`
3. Deploy. From then on, every `git push` updates the live site
   automatically.

Either option works fine together with Firebase for the data/login side —
GitHub/Vercel/Netlify just serves the app's files; Firebase stores the
data and handles logins.

## Notes

- The Firebase web config in `src/firebase.js` is safe to commit — it's a
  client identifier, not a secret key. Real protection comes from the
  Firestore rules and Firebase Authentication.
- Editing a staff member's username/password isn't available from the
  Settings screen — remove and re-add the user if that's needed. Full
  self-service password reset would need a real email address per user or
  a small server-side function; ask if you want that added later.
