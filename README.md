### Outlook Inbox Attachments Downloader 

# Monitor and download attachments in unread mails from outlook app in Windows.

```Inspiration - Automate the boring stuff with Python```

Runs outlook app, continously monitor for new mails and download attachments to the default path and set the unread mails in inbox to read. Works in windows with outlook app.

### Install: 
```bash
git clone "https://github.com/itsmepvr/outlook-attachment-downloader.git"
cd outlook-attachment-downloader
```
Install pipenv for environment (optional)
```bash
pip install -r requirements.txt
```

### Usage:
```bash
python index.py
```
* Change the default path to download attachments in index.py file at user_path
* Put the mail subject which you want to download for incoming mails or keep it empty('') to download for any mails.

## Built With

* Python 3.6
* Pipenv

## licensed

This project is licensed under the MIT License

## Authors

* **Venkata Ramana P** - [itsmepvr](https://itsmepvr.github.io)