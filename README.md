# File on Send — setup guide

Choose where each sent message gets filed, right from the compose window in New Outlook for Mac.

## How it works

1. **The add-in (compose window).** Click **File to Folder** on the toolbar, then search your folders and pick one. The add-in tags the draft with a category such as `File: Courses/CHEG 345/Exams`. You can also type a path that doesn't exist yet, and it will be created.
2. **The Smart Alert (when you click Send).** If you haven't picked a folder, Outlook asks you:
   - **Choose folder** opens the picker.
   - **Send Anyway** sends normally, and the message stays in Sent Items.
3. **"File sent mail" flow (Power Automate).** When the message lands in Sent Items (after your 30-second delay), the flow reads the tag, moves the message to the folder (creating the folder if needed), and removes the tag. If anything goes wrong, the message stays in Sent Items with a **Filing failed** category, so nothing is lost.
4. **"Export folder list" flow (Power Automate).** Each morning it walks your whole folder tree and emails you the current list. Paste it into the add-in with **Update list** to stay current. Folders you create from the picker are added to the list immediately, so you only need to refresh for folders made elsewhere, or for renames and deletions.

Your folder list is stored in your mailbox (the add-in's settings), not on the web. The only things hosted publicly are the add-in's code and icons.

---

## Part 1 — Host the add-in files (GitHub Pages, free, ~10 min)

Outlook loads add-in pages from an HTTPS web address, and GitHub Pages provides one for free.

1. Create a free account at github.com (any email works).
2. Click **+ → New repository**, name it `file-on-send`, set it to **Public**, and click **Create repository**.
3. Click **uploading an existing file**. Drag in everything from this folder (`taskpane.html`, `taskpane.js`, `commands.html`, `launchevent.js`, `manifest.xml`, `README.md`, and the whole `assets` folder), then click **Commit changes**.
4. Go to **Settings → Pages**. Under *Build and deployment*, choose **Deploy from a branch**, then **main** and **/ (root)**, and click **Save**.
5. After a minute or two, open `https://mblaustein78.github.io/file-on-send/taskpane.html`. You should see the picker with a yellow "Preview mode" banner.
6. `manifest.xml` is already set up for the address `https://mblaustein78.github.io/file-on-send/`, so keep the repository name exactly `file-on-send`.

## Part 2 — Build the "File sent mail" flow

Go to make.powerautomate.com, signed in with your UD account. Choose **Create → Automated cloud flow**, name it **File sent mail**, pick the trigger **When a new email arrives (V3)** (Office 365 Outlook), and click **Create**.

Wherever a step below says **Send an HTTP request**, use the *Office 365 Outlook* action of that name. It's a standard action, so no premium license is needed. Rename each action exactly as shown (click the action title), because later expressions refer to those names. Text shown as `@{…}` is an **expression**: put the cursor in the field, choose the *fx* / expression option, and paste everything between `@{` and `}`.

| # | Action | Settings |
|---|---|---|
| T | **When a new email arrives (V3)** | Folder: **Sent Items**. Include Attachments: **No** |
| 1 | **Initialize variable** | Name `folderId`, Type String, Value `msgfolderroot` |
| 2 | **Initialize variable** | Name `path`, Type String, Value *(leave empty)* |
| 3 | **Send an HTTP request**, rename **Get categories** | Method GET. URI `https://graph.microsoft.com/v1.0/me/messages/@{triggerOutputs()?['body/id']}?$select=categories` |
| 4 | **Filter array**, rename **File tag** | From `@{body('Get_categories')?['categories']}`. Edit in advanced mode: `@startsWith(item(), 'File: ')` |
| 5 | **Filter array**, rename **Other categories** | From the same as step 4. Advanced mode: `@not(startsWith(item(), 'File: '))` |
| 6 | **Condition** | `@{length(body('File_tag'))}` *is greater than* `0` |

In the **True** branch of the Condition:

| # | Action | Settings |
|---|---|---|
| 7 | **Set variable** | `path` = `@{substring(first(body('File_tag')), 6)}` |
| 8 | **Scope**, rename **Try** | *(the actions below go inside it)* |

Inside **Try**:

| # | Action | Settings |
|---|---|---|
| 8a | **Apply to each**, rename **Each folder name** | Input `@{split(variables('path'), '/')}`. In the action's **Settings**, turn on *Concurrency control* and set the degree of parallelism to **1** |
| 8b | ↳ **Send an HTTP request**, rename **Find folder** | GET `https://graph.microsoft.com/v1.0/me/mailFolders/@{variables('folderId')}/childFolders?$filter=displayName eq '@{encodeUriComponent(replace(items('Each_folder_name'), '''', ''''''))}'&$select=id` |
| 8c | ↳ **Condition** | `@{length(body('Find_folder')?['value'])}` *is greater than* `0` |
| | ↳ True: **Set variable** | `folderId` = `@{first(body('Find_folder')?['value'])?['id']}` |
| | ↳ False: **Send an HTTP request**, rename **Create folder** | POST `https://graph.microsoft.com/v1.0/me/mailFolders/@{variables('folderId')}/childFolders`. Body `{"displayName": "@{items('Each_folder_name')}"}`. Content-Type `application/json` |
| | ↳ False: **Set variable** | `folderId` = `@{body('Create_folder')?['id']}` |
| 8d | **Send an HTTP request**, rename **Remove tag** *(after the loop, still inside Try)* | PATCH `https://graph.microsoft.com/v1.0/me/messages/@{triggerOutputs()?['body/id']}`. Body `{"categories": @{body('Other_categories')}}`. Content-Type `application/json` |
| 8e | **Send an HTTP request**, rename **Move message** | POST `https://graph.microsoft.com/v1.0/me/messages/@{triggerOutputs()?['body/id']}/move`. Body `{"destinationId": "@{variables('folderId')}"}`. Content-Type `application/json` |

Still in the **True** branch, after **Try**:

| # | Action | Settings |
|---|---|---|
| 9 | **Scope**, rename **If filing failed** | Open its **Settings → Run after**, *uncheck* "is successful", and *check* "has failed" and "has timed out" for **Try** |
| 9a | ↳ **Send an HTTP request**, rename **Flag failure** | PATCH `https://graph.microsoft.com/v1.0/me/messages/@{triggerOutputs()?['body/id']}`. Body `{"categories": @{union(body('Other_categories'), createArray('Filing failed'))}}`. Content-Type `application/json` |

Leave the **False** branch of step 6 empty. Untagged messages simply stay in Sent Items. Save the flow.

**Optional tidy-up.** Each folder you pick adds a `File: …` entry to your Outlook category list. If those pile up, you can delete them in Outlook's category settings at any time. The add-in recreates any it needs.

## Part 3 — Build the "Export folder list" flow

Choose **Create → Scheduled cloud flow**, name it **Export folder list**, set it to repeat every **1 day** at a morning time, and click **Create**.

| # | Action | Settings |
|---|---|---|
| 1 | **Initialize variable** | Name `queue`, Type Array, Value `[{"id":"msgfolderroot","path":""}]` |
| 2 | **Initialize variable** | Name `paths`, Type Array, Value `[]` |
| 3 | **Initialize variable** | Name `current`, Type Object, Value `{}` |
| 4 | **Do until** | Loop until `@{length(variables('queue'))}` *is equal to* `0`. Under *Change limits*, set Count `2000` and Timeout `PT1H` |

Inside **Do until**:

| # | Action | Settings |
|---|---|---|
| 4a | **Set variable** | `current` = `@{first(variables('queue'))}` |
| 4b | **Compose**, rename **Rest of queue** | `@{skip(variables('queue'), 1)}` |
| 4c | **Set variable** | `queue` = `@{outputs('Rest_of_queue')}` |
| 4d | **Send an HTTP request**, rename **List children** | GET `https://graph.microsoft.com/v1.0/me/mailFolders/@{variables('current')?['id']}/childFolders?$top=250&$select=id,displayName,childFolderCount` |
| 4e | **Apply to each**, rename **Each child** | Input `@{body('List_children')?['value']}`. Concurrency control **On**, degree **1** |
| | ↳ **Condition**, rename **Not a system folder** | Advanced mode: `@not(contains(createArray('Deleted Items','Junk Email','Outbox','Drafts','Conversation History','Sync Issues','RSS Feeds','RSS Subscriptions'), items('Each_child')?['displayName']))` |
| | ↳ True: **Compose**, rename **Child path** | `@{if(empty(variables('current')?['path']), items('Each_child')?['displayName'], concat(variables('current')?['path'], '/', items('Each_child')?['displayName']))}` |
| | ↳ True: **Append to array variable** | `paths`, value `@{outputs('Child_path')}` |
| | ↳ True: **Condition** | `@{items('Each_child')?['childFolderCount']}` *is greater than* `0` |
| | ↳ ↳ True: **Append to array variable** | `queue`, value (expression) `addProperty(addProperty(json('{}'), 'id', items('Each_child')?['id']), 'path', outputs('Child_path'))` |

After **Do until**:

| # | Action | Settings |
|---|---|---|
| 5 | **Compose**, rename **Folder list** | `@{join(sort(variables('paths')), decodeUriComponent('%0A'))}` |
| 6 | **Send an email (V2)** | To: your UD address. Subject `File on Send – folder list`. Body (switch the editor to code view `</>`): `<pre>@{outputs('Folder_list')}</pre>` |

Save, then click **Run** once to get your first list right away.

## Part 4 — Install the add-in in Outlook

1. In New Outlook for Mac, open the **Apps** / **Get Add-ins** button, then **My add-ins**.
2. Under **Custom Add-ins**, choose **Add a custom add-in → Add from file…** and pick your edited `manifest.xml`. Accept the warning.
3. Open a new message. A **File to Folder** button appears on the message toolbar (it may be under the **…** / Apps menu).
4. Click it, click **Update list**, paste the list from the "Export folder list" email, and click **Save list**.

The add-in installs into your mailbox, so it also appears in Outlook on the web.

## Part 5 — Test

1. Write a message to yourself and click **File to Folder**. Pick a test folder (or type a new path such as `Test/File on Send`) and click **Send**.
2. Write another message and send it *without* picking a folder. The prompt should appear. Try **Choose folder** and **Send Anyway**.
3. After the 30-second delay plus up to a minute for the flow, check that the first message is in the chosen folder and the second is in Sent Items.
4. If a message doesn't move, open the **File sent mail** flow's run history in Power Automate. The failing step will be marked in red.

## Keeping it current

- **New folder while working:** type its full path in the picker. It's created when the message is filed and added to your list right away.
- **Folders created, renamed, or deleted elsewhere:** paste the latest list from the daily email (**Update list → paste → Save list**). It takes about ten seconds whenever you notice something missing.

## Good to know

- **Every send is checked**, including replies. If you don't want a message filed, click **Send Anyway**.
- **Folder names containing `/`** can't be used, because `/` separates folder levels.
- **Outlook mobile** doesn't run this add-in. Mail sent from your phone stays in Sent Items.
- **If the add-in can't run** (for example, when you're offline), Outlook sends normally and the message stays in Sent Items.
- **The list holds several hundred folders.** If **Save list** ever fails for size, remove folders you never file to.
- **Updating the add-in later:** upload changed files to the GitHub repository. Outlook picks them up the next time the pane opens, so you don't need to reinstall unless `manifest.xml` changes.
