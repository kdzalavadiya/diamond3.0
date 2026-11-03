DIAMOND LEDGER - NETLIFY VERSION  (no Apps Script, no server)
================================================================
Files: index.html, app.js, Code.js, gas-shim.js, config.js
Only config.js needs editing (put your Google Client ID in it).

STEP 1 - Google Cloud (one time, free)
 1. Go to https://console.cloud.google.com and create a new project.
 2. APIs & Services > Library > search "Google Sheets API" > Enable.
 3. APIs & Services > OAuth consent screen (Google Auth Platform):
      - User type: External. Fill app name + your email.
      - Data access / Scopes: add  .../auth/spreadsheets  and  .../auth/userinfo.email
      - Audience / Test users: ADD THE GMAIL OF EVERY PERSON WHO WILL USE THE APP (up to 100).
 4. APIs & Services > Credentials > Create credentials > OAuth client ID
      - Application type: Web application
      - Authorized JavaScript origins: your Netlify URL (see Step 2), e.g. https://my-ledger.netlify.app
        (Redirect URIs: leave empty)
 5. Copy the Client ID and paste it into config.js.

STEP 2 - Netlify
 1. Unzip this folder. Go to https://app.netlify.com/drop and drag the folder in.
 2. Copy your site URL (you can rename it under Site configuration > Change site name).
 3. Add that URL to "Authorized JavaScript origins" (Step 1.4).
 4. After editing config.js with your Client ID, upload the folder again
    (Deploys tab > drag the folder onto the page).

STEP 3 - Daily use
 Owner: share the Google Sheet with each employee's Gmail as Editor (can change data)
        or Viewer (read only). Change it any time in the Sheet's Share menu.
 Employee: open the Netlify link > Sign in with Google (click Advanced > Continue if Google
        shows "app not verified") > click "Connect Sheet" > paste the sheet link.
        The Purchases and Sales tabs are created automatically if missing.
