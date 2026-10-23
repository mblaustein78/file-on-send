# File on Send: reference guide

File each sent message into a folder you choose, instead of leaving it in Sent Items.
Built for New Outlook for Mac on the University of Delaware Microsoft 365 account. It also works in Outlook on the web and Outlook for Windows.

*Last updated October 2026. This guide describes the design as actually built and tested.*

---

## How it works

| Piece | Where it lives | What it does |
|---|---|---|
| **File on Send add-in** | Installed in your mailbox; code hosted on GitHub Pages | A **File to Folder** button in the compose window opens a searchable folder picker. Choosing a folder tags the draft with a category such as `File: Courses/CHEG 345/Exams`. |
| **"File sent mail" flow** | Power Automate (cloud) | Every 5 minutes it checks your 50 most recent Sent Items. Each message tagged `File: …` is moved to that folder, and any missing folders are created first. Afterwards it removes the `File: …` label from the message and the matching entry from your category list, so both stay tidy. |
| **"Export folder list" flow** | Power Automate (cloud) | Each morning it emails you a list of all your folders, one full path per line. Paste the list into the picker to keep it current. |

**Day to day:**
- Before sending a message you want filed, click **File to Folder**. Search for and pick a folder, or type a new path such as `Research/Grants/2027`, which will be created.
- If you don't pick a folder, the message stays in Sent Items as usual.
- Once a message is filed, its `File: …` label is removed. Your other categories stay.
- Allow up to about 5 minutes after sending for the move. Your 30-second send delay is fine.

---

## Part 1: Add-in hosting (GitHub Pages)

- Repository: `github.com/mblaustein78/file-on-send` (public)
- Live address: `https://mblaustein78.github.io/file-on-send/`
- Pages setting: **Settings → Pages → Deploy from a branch → main / (root)**

**Files in the repository**

| File | Purpose |
|---|---|
| `manifest.xml` | Add-in definition, **version 1.1** (no Send prompt). You install this into Outlook; it isn't loaded from the website. |
| `taskpane.html`, `taskpane.js` | The folder picker, **v1.2**. The version shows at the bottom of the pane. |
| `commands.html`, `launchevent.js` | Left over from the old Send prompt. Unused, but harmless to keep. |
| `assets/icon-*.png` | Button icons |

**To update the add-in code:** upload the changed files with **Add file → Upload files → Commit changes**, wait about 2 minutes, then quit and reopen Outlook. Only `manifest.xml` changes require reinstalling the add-in.

---

## Part 2: Install the add-in (once per mailbox)

1. In Outlook, go to **Apps → Add apps → My add-ins → Custom add-ins → Add a custom add-in → Add from file…**
2. Choose `manifest.xml` (version 1.1) and click **Install**.
3. Quit Outlook (⌘Q) and reopen it.
4. Open a new message. **File to Folder** appears on the toolbar, or under **Apps** / **⋯**.

The add-in installs into your **mailbox**, not into one computer. See Part 6.

**First use:** click **File to Folder → Update list**, paste the folder list from the "Export folder list" email, and click **Save list**.

---

## Part 3: "File sent mail" flow (as built)

**Power Automate tips, learned the hard way:**
- **Formulas** (shown below as `code`) must be entered with **fx** (or **/** → *Insert expression*). The result must appear as a chip such as **concat(…)**, not as typed text.
- Don't pick look-alike items from the dynamic-content list in place of a formula. They may point at the wrong data.
- Formulas refer to steps by name, with spaces turned into underscores (*Get Categories* → `Get_Categories`). Keep the step names exactly as listed.
- In conditions, enter both sides as formulas (for example `greater(…)` **is equal to** `true`). A typed `0` is treated as text.
- A condition with an empty extra row won't save. Delete the row with **⋯ → Delete**.
- Inside web addresses, curly braces must be written as `%7B` and `%7D`.
- Use the **Office 365 Outlook → Send an HTTP request** action, not the premium "HTTP" action.

**Structure** (indentation shows what sits inside what):

```
Recurrence ............................ every 5 minutes
Initialize variable ................... folderId (String) = msgfolderroot
Initialize variable 1 ................. path (String) = (empty)
Get Categories ........................ GET recent Sent Items
File tag .............................. Filter array: only messages tagged "File:"
Each message .......................... Apply to each (concurrency 1)
    Set variable 3 .................... folderId = msgfolderroot
    Target path ....................... Compose: folder path from the tag
    Has tag ........................... Condition
      True:
        Set variable .................. path = Target path
        Each folder name .............. Apply to each (concurrency 1)
            Find folder ............... GET child folder by name
            Found it .................. Condition
              True:  Set variable 1 ... folderId = found folder
              False: Create folder .... POST new folder
                     Set variable 2 ... folderId = new folder
        Move message .................. POST move
        Keep categories ............... Filter array: the message's other categories
        Remove tag .................... PATCH the moved message's categories
        Get category list ............. GET master categories      (cleanup)
        This category ................. Filter array               (cleanup)
        Each category ................. Apply to each              (cleanup)
            Delete category ........... DELETE                     (cleanup)
```

**Step details**

| Step | Settings |
|---|---|
| **Recurrence** | Interval `5`, Frequency **Minute** |
| **Get Categories** | GET. URI (typed, no fx): `https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages?$top=50&$select=id,categories&$orderby=sentDateTime%20desc` |
| **File tag** | From: `body('Get_Categories')?['value']`. Row: `join(coalesce(item()?['categories'], json('[]')), '\|')` **contains** `File:` |
| **Each message** | Input: `body('File_tag')`. Settings → Concurrency **On**, **1** |
| **Set variable 3** | `folderId` = `msgfolderroot` (typed) |
| **Target path** | `trim(first(split(last(split(join(items('Each_message')?['categories'], '\|'), 'File:')), '\|')))` |
| **Has tag** | `greater(length(outputs('Target_path')), 0)` **is equal to** `true` |
| **Set variable** | `path` = `outputs('Target_path')` |
| **Each folder name** | Input: `split(variables('path'), '/')`. Concurrency **On**, **1** |
| **Find folder** | GET. URI: `concat('https://graph.microsoft.com/v1.0/me/mailFolders/', variables('folderId'), '/childFolders?$filter=displayName%20eq%20''', encodeUriComponent(replace(items('Each_folder_name'), '''', '''''')), '''&$select=id')` |
| **Found it** | `length(body('Find_folder')?['value'])` **is greater than** `0` |
| **Set variable 1** | `folderId` = `first(body('Find_folder')?['value'])?['id']` |
| **Create folder** | POST. URI: `concat('https://graph.microsoft.com/v1.0/me/mailFolders/', variables('folderId'), '/childFolders')`. Body: `setProperty(json('{}'), 'displayName', items('Each_folder_name'))`. Content-Type `application/json` |
| **Set variable 2** | `folderId` = `body('Create_folder')?['id']` |
| **Move message** | POST. URI: `concat('https://graph.microsoft.com/v1.0/me/messages/', items('Each_message')?['id'], '/move')`. Body: `setProperty(json('{}'), 'destinationId', variables('folderId'))`. Content-Type `application/json` |
| **Keep categories** | From: `coalesce(items('Each_message')?['categories'], json('[]'))`. Row: `item()` **does not start with** `File:` |
| **Remove tag** | PATCH. URI: `concat('https://graph.microsoft.com/v1.0/me/messages/', body('Move_message')?['id'])` (use the move step's exact name; a moved message gets a new ID). Body: `setProperty(json('{}'), 'categories', body('Keep_categories'))`. Content-Type `application/json` |
| **Get category list** | GET. URI (typed): `https://graph.microsoft.com/v1.0/me/outlook/masterCategories` |
| **This category** | From: `body('Get_category_list')?['value']`. Row: `toLower(item()?['displayName'])` **is equal to** `toLower(concat('File: ', variables('path')))` |
| **Each category** | Input: `body('This_category')` |
| **Delete category** | DELETE. URI: `concat('https://graph.microsoft.com/v1.0/me/outlook/masterCategories/', items('Each_category')?['id'])` |

(In the table, `\|` stands for a plain vertical bar `|`.)

**Why it's built this way:**
- **Polling every 5 minutes instead of "When a new email arrives":** that trigger missed or delayed Sent Items messages, especially after the flow had been edited.
- **Categories as the marker:** they survive sending in New Outlook for Mac, although they appear on the sent copy a little after it lands. The add-in also stores a hidden custom property, but the flow doesn't rely on it.
- **The tag is removed after the move, not before:** if a move fails, the message stays in Sent Items with its tag and is retried on the next 5-minute check.
- **Deleting the category-list entry is safe:** it doesn't remove the label from any message, and the add-in recreates the entry the next time you pick that folder.

---

## Part 4: "Export folder list" flow (as built)

```
Recurrence ............................ every 1 day
Initialize variable ................... queue (Array) = [{"id":"msgfolderroot","path":""}]
Initialize variable 1 ................. paths (Array) = []
Initialize variable 2 ................. current (Object) = {}
Do until .............................. length(variables('queue')) is equal to 0   (Count 2000, Timeout PT1H)
    Set variable ...................... current = first(variables('queue'))
    Rest of queue ..................... Compose: skip(variables('queue'), 1)
    Set variable 1 .................... queue = outputs('Rest_of_queue')
    List children ..................... GET (URI below)
    Each child ........................ Apply to each body('List_children')?['value']  (concurrency 1)
        Not a system folder ........... Condition (below)
          True:
            Child path ................ Compose (below)
            append to array variable .. paths += outputs('Child_path')
            Condition 1 ............... greater(items('Each_child')?['childFolderCount'], 0) is equal to true
              True: Append to array variable 1 ... queue += (below)
Send an email (V2) .................... To: you. Body: join(sort(variables('paths')), '<br>')
```

| Step | Formula |
|---|---|
| **List children** URI | `concat('https://graph.microsoft.com/v1.0/me/mailFolders/', variables('current')?['id'], '/childFolders?$top=250&$select=id,displayName,childFolderCount')` |
| **Not a system folder** | `contains(createArray('Deleted Items','Junk Email','Outbox','Drafts','Conversation History','Sync Issues','RSS Feeds','RSS Subscriptions','Sent Items'), items('Each_child')?['displayName'])` **is equal to** `false` |
| **Child path** | `if(empty(variables('current')?['path']), items('Each_child')?['displayName'], concat(variables('current')?['path'], '/', items('Each_child')?['displayName']))` |
| **Append to array variable 1** | `addProperty(addProperty(json('{}'), 'id', items('Each_child')?['id']), 'path', outputs('Child_path'))` |

Every step from *Set variable* to *Each child* must sit **inside** the Do until box.

---

## Part 5: Keeping things current

- **New folder while writing:** type its full path in the picker. It's created when the message is filed and added to your list right away.
- **Folders created, renamed or deleted elsewhere:** **File to Folder → Update list**, paste the latest daily email, **Save list**.
- **Removing a bad entry from the picker:** **Update list**, delete the line, **Save list**.
- **Duplicate subfolder names** (e.g. `Students/Advising` and `Research/Advising`) are fine, because everything uses full paths.
- **Avoid `/` in folder names.** The slash separates levels.

---

## Part 6: Using a second computer

Nothing needs reinstalling:
- **The add-in** is installed in your **mailbox**, so it appears in New Outlook on any computer signed in to `mblau@udel.edu`, and in Outlook on the web.
- **The folder list and recent choices** are stored in your mailbox (the add-in's settings), so they follow you.
- **The flows** run in Microsoft's cloud, so they work whichever computer you send from, including your phone. Mail sent from Outlook mobile can't be tagged, though, because the add-in doesn't run on mobile.

**On the other computer:**
1. Quit and reopen Outlook.
2. Open a new message and check for **File to Folder** (toolbar, or **Apps** / **⋯**).
3. Open it and check the bottom of the pane: it should show your folder count and **v1.2**.
4. If an old "File on Send is working / taking longer" popup appears when you send, that computer is holding a cached copy of the old add-in. Quit Outlook fully and reopen it. If it persists: **Apps → Add apps → My add-ins**, remove File on Send, and add `manifest.xml` (v1.1) again from file.
5. If **File to Folder** doesn't appear at all after a restart, install it as in Part 2.

If the other computer uses classic Outlook for Windows, the add-in works there too.

---

## Part 7: Setting this up for a colleague

Nothing in the code is tied to your account. The flows act on whoever runs them, and each person's folder list stays in their own mailbox.

**1. The add-in.** Give them `manifest.xml` (v1.1). They install it as in Part 2.
- It loads the code from **your** GitHub Pages site, so they depend on your repository staying in place. Any change you upload affects them too.
- **To be independent:** they create their own copy of the repository (**Fork** on GitHub), turn on Pages, and replace the six `mblaustein78.github.io` addresses in `manifest.xml` with their own. That's the only code change ever needed.

**2. The flows.** Copy both rather than rebuilding them:
- Open the flow, then **⋯ → Send a copy**, enter their UD email, and click **Send**. Do this for **File sent mail** and **Export folder list**.
- They open each copy, sign in with their own Outlook connection when prompted, then **Save** and turn it **On**.
- If *Send a copy* isn't available: **Export → Package (.zip)**. They use **Import package** in Power Automate and choose their own connection.
- **One edit:** in **Export folder list → Send an email (V2)**, change **To** to their address.

**3. First run.** They run **Export folder list** once (**Test → Manually**), then paste the emailed list into the picker (**File to Folder → Update list → Save list**).

**Requirements:** New Outlook for Mac, Outlook on the web, or classic Outlook for Windows; Power Automate access; permission to install custom add-ins. UD's Microsoft 365 setup allows all three.

---

## Part 8: Backups

- **Add-in files:** the GitHub repository is the master copy. Keep a downloaded copy of the repository as well (**Code → Download ZIP**).
- **Flows:** after any change, export each one (**Export → Package (.zip)**) and keep the packages with the add-in files. A package can be imported to restore a flow exactly.

---

## Part 9: Troubleshooting

| Symptom | Check |
|---|---|
| A message wasn't filed | Is it still in Sent Items with a `File: …` label? Open **File sent mail → Run history** and look at the latest run's red step. |
| Run history shows no runs | On the flow's details page, **Status** should be **On**. |
| The folder picker is empty | **Update list** → paste the folder-list email. |
| The picker shows an old version | Quit and reopen Outlook. If it still shows the old version, check that GitHub has the latest `taskpane.html`/`taskpane.js`. |
| A run is stuck for minutes | A step is retrying an error. Open the step, then **Settings → Retry policy → None** to see the error right away. |
