import { initializeApp, getApps, getApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

// 1. Go to https://console.firebase.google.com -> create a project
// 2. Add a "Web App" inside that project -> it gives you this config object
// 3. Paste your real values in below (these are safe to commit/publish —
//    they are not secret keys, access is controlled by Firestore/Auth rules)
const firebaseConfig = {
  apiKey: "AIzaSyBEGPLYQrzcLDZxNMmxfVDZV0aEbMc4QU4",
  authDomain: "twig-global-erp.firebaseapp.com",
  projectId: "twig-global-erp",
  storageBucket: "twig-global-erp.firebasestorage.app",
  messagingSenderId: "451889498454",
  appId: "1:451889498454:web:5be7ca74a3f27a1be2eef7",
};

export const app = getApps().length ? getApp() : initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);

// Used only when the admin creates a new staff account, so that creating
// the account doesn't log the admin out of their own session.
export function getSecondaryAuth() {
  const secondaryApp = initializeApp(firebaseConfig, "secondary-" + Date.now());
  return { authInstance: getAuth(secondaryApp), appInstance: secondaryApp };
}

// Staff sign in with a plain username, but Firebase Auth needs an email
// shape, so we turn "john" into "john@hardwareerp.local" behind the scenes.
export const toAuthEmail = (username) => `${username.trim().toLowerCase()}@hardwareerp.local`;
