
# teams-chats-export

Script to export your Microsoft Teams [chats](https://learn.microsoft.com/en-us/graph/api/resources/chat?view=graph-rest-1.0)
and view them as .html files. Chats are conversations separate from those on Teams channels. These include one-on-one chats,
group chats, and meeting chats.

Unlike similar export tools, this does NOT require using Windows or PowerShell, or registering
an application with the Microsoft Identity Platform. The only requirement is Python.

## Instructions

To archive your chats:

```sh
# set up and activate a virtual environment
python -m venv teams-chats-export-env
. ./teams-chats-export-env/bin/activate

# install dependencies
pip install -r requirements.txt

# run two steps to download and generate html files in the './archive' directory
# to use another dir, set --output-dir=/some/other/dir
python teams_chats_export.py download
python teams_chats_export.py generate_html
```

The download step will open a browser window for you to authenticate.

Files will be written to directories named `data` and `html` within the output directory.

You can run the download step on your archive directory periodically to update the data.
Previously downloaded data, including anything deleted on Teams since the last time you
ran the script, will be retained in the directory. (If you want to be really safe, consider
making a backup copy of the archive directory from time to time.)

Only "new" messages are downloaded, but this logic may not catch certain edge cases.
If you're missing messages or you don't see edits, or you just want to be absolutely sure
you're downloading everything, run `python teams_chats_export.py download --force`
to re-process all the available messages on Teams.

If you want to customize the html outputs, edit `templates/chat.jinja` to your
liking and re-run the generate_html step above.

## Limitations

- Doesn't render reactions in the .html

- Not all attachment types are handled properly. (Meta)data for unhandled attachments
  is written directly to the .html file with the string prefix "Attachment (raw data):"
