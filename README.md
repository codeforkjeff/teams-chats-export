
# teams-chats-export

Script to export your Microsoft Teams [chats](https://learn.microsoft.com/en-us/graph/api/resources/chat?view=graph-rest-1.0)
and view them as .html files. Chats are conversations separate from those on Teams channels. These include one-on-one chats,
group chats, and meeting chats.

## Design Goals

Unlike similar export tools, this does NOT require using Windows or PowerShell. The only requirement is Python.

Originally you did not need to register an application with the Microsoft Identity Platform or Microsoft Entra ID,
but as of Sept 9th, 2024, [that is no longer the case](https://pnp.github.io/blog/post/changes-pnp-management-shell-registration/).

This was created specifically in response to a new retention policy for Teams chats
implemented at my workplace.

## Instructions


### Create an Entra ID application if needed

You'll need either an existing Entra ID application with permissions to access Teams,
or you'll need to create a new one.

If you have an existing application, find its application/client ID.

If you need to create one, the steps are [here](https://pnp.github.io/powershell/articles/registerapplication.html#manually-create-an-app-registration-for-interactive-login).
In a nutshell:

- Log into https://entra.microsoft.com/
- Navigate to Applications -> App Registrations
- Click "New registration" and give it a name like "Teams API"
- Navigate to Authentication for the new app
- Click "Add a platform" and select "Mobile and desktop applications"
- Enter the following custom redirect URI: http://localhost:8400
- Save your changes

You don't need to configure specific API permissions. The default "User.Read" delegated
permission will give users access to their own Teams data.

Copy your new application's "Application (client) ID" from the Overview screen. It should look
something like "31359c7f-bd7e-475c-86db-fdb8c937548e" (that's the ID for the old built-in
PnP Management Shell Entra App that Microsoft has disabled, it won't work).

### Install the script

Do this once:

```sh
# set up and activate a virtual environment
python -m venv teams-chats-export-env
. ./teams-chats-export-env/bin/activate

# install dependencies
pip install -r requirements.txt
```

### Run the script

To archive your chats:

```sh
# activate the virtual env every time you want to use it
. ./teams-chats-export-env/bin/activate

# run two steps to download and generate html files in the './archive' directory
# to use another dir, set --output-dir=/some/other/dir

# step one: download
# replace value for CLIENT_ID with your own ID
CLIENT_ID="31359c7f-bd7e-475c-86db-fdb8c937548e" python teams_chats_export.py download

# step two: generate html files
python teams_chats_export.py generate_html
```

The download step will open a browser window for you to authenticate.

Files will be written to directories named `data` and `html` within the output directory.
The `html` directory includes an `index.html` containing a listing of the chats.

You can run the download step on your archive directory periodically to update the data.
Previously downloaded data, including anything deleted on Teams since the last time you
ran the script, will be retained in the directory. (If you want to be really safe, consider
making a backup copy of the archive directory from time to time.)

Only "new" messages are downloaded, but this logic may not catch certain edge cases.
If you're missing messages or you don't see edits, or you just want to be absolutely sure
you're downloading everything, run `python teams_chats_export.py download --force`
to re-process all the available messages on Teams.

If you want to customize the html outputs, edit the files in `templates/` to your
liking and re-run the generate_html step above.

## Limitations

- Doesn't render reactions in the .html

- Not all attachment types are handled properly. (Meta)data for unhandled attachments
  is written directly to the .html file with the string prefix "Attachment (raw data):"
