


from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools


def upload_google_api(file_name, file_path, folder_id):
    try :
        import argparse
        flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
    except ImportError:
        flags = None

    SCOPES = 'https://www.googleapis.com/auth/drive.file'
    store = file.Storage('storage.json')
    creds = store.get()

    if not creds or creds.invalid:
        print("make new storage data file ")
        flow = client.flow_from_clientsecrets('client_secret_drive.json', SCOPES)
        creds = tools.run_flow(flow, store, flags) \
                if flags else tools.run(flow, store)

    DRIVE = build('drive', 'v3', http=creds.authorize(Http()))

    FILES = ((file_name),)
    folder_id = folder_id  # 폴더코드

    for file_title in FILES :
        file_name = file_title
        metadata = {'name': file_name, 'parents' : [folder_id],
                    'mimeType': None
                    }
        media = str(file_path + "\\" + file_name)
        print(media)
        res = DRIVE.files().create(body=metadata, media_body=media).execute()
        if res:
            print('Uploaded "%s" (%s)' % (file_name, res['mimeType']))


if __name__ == '__main__':
    upload_google_api('hello.txt', "C:\Projects\quant\data", '1DDMvvmZmYhXyvX2mBZ5F2SE2t5IJr3DE')

    # from Upload_google_drive import upload_google_api
    # import후에
    # upload_google_api(file_name, file_path, folder_id)
    # 위와같이 입력하여 사용
