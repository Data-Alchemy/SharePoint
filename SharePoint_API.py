from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.runtime.client_request_exception import ClientRequestException
import os
import json


'''
This class builds off of the office 365 class to facilitate programmatic operations on sharepoint
'''

class SharePoint_Manager():

    def __init__(self,username,password,base_url):
        self.username = username
        self.password = password
        self.base_url = base_url
        user_credentials = UserCredential(username,password)
        self.ctx = ClientContext(base_url).with_credentials(user_credentials)
        self.web = self.ctx.web
        self.ctx.load(self.web)
        self.ctx.execute_query()

    @property
    def validate_parms(self):
        return {'username': self.username,
                'password': self.password,
                'base_url': self.base_url,
                }

    def check_for_folder(self,ck_path:str)->bool:
        try:
            self.web.get_folder_by_server_relative_url(ck_path).get().execute_query()
            return True
        except ClientRequestException as e:
            if e.response.status_code == 404:
                return False
            else:
                raise ValueError(e.response.text)


    def get_sharepoint_folders(self,folder_path:str)->list:
        lib = self.web.get_folder_by_server_relative_url(folder_path)
        folders = lib.folders
        self.ctx.load(folders)
        self.ctx.execute_query()
        folder_list = []
        for myfolder in folders:
            folder_list.append(myfolder.properties["ServerRelativeUrl"])
        return folder_list


    def get_sharepoint_files(self,folder_path: str) -> list:
        lib = self.web.get_folder_by_server_relative_url(folder_path)
        files = lib.files
        self.ctx.load(files)
        self.ctx.execute_query()
        file_list = []
        for myfile in files:
            file_list.append(myfile.properties["ServerRelativeUrl"])
        return file_list

    def upload_file_to_sharepoint_title(self,source_path:str,target_path:str):
        file_data        = open(source_path,'rb').read()
        file_name        = os.path.basename(source_path)
        print(file_name)
        target      = self.ctx.web.lists.get_by_title(target_path).root_folder
        upload_file = target.upload_file(file_name,file_data).execute_query()
        print(f'file uploaded to path: {upload_file.serverRelativeUrl}')

    def create_folder(self,path)->str:
        creation_list = []
        try:
            folder_paths = path.split('/')
            i = 0
            while i <= len(folder_paths):
                check_path = '/'.join(folder_paths[0:i])

                if self.check_for_folder(check_path) != True:
                    flder = self.ctx.web.folders.add(check_path)
                    mkdir = self.ctx.load(flder).execute_query()
                    creation_list.append(mkdir.serverRelativeUrl)
                i += 1
            return None#f'Created path(s) :{creation_list}'
        except Exception as e:
           return f'Unable to create folder {e}'




    def upload_file_to_sharepoint_path(self,source_path:str,target_path:str):
        target_path = f'SSRS Reports{target_path}'
        #check if folder exists if not create #
        self.create_folder(target_path)

        file_data        = open(source_path,'rb').read()
        file_name        = os.path.basename(source_path)

        try:
            target      = self.ctx.web.get_folder_by_server_relative_url(target_path)
            upload_file = target.upload_file(file_name,file_data).execute_query()
            return f'file uploaded to path: {upload_file.serverRelativeUrl}'
        except Exception as e:
            return f'Error upload of source file {source_path} failed \n error is {e}'


'''
# sample usage #
SharePoint_Manager(username='username',password=pwd',base_url='https://seapeakonline.sharepoint.com/sites/Reports/').check_for_folder('Delta')
'''



