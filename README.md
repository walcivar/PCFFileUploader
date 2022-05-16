If you want, you can download the code and make some changes before uploading it to your environment, such as changing the solution prefix.
But if you want to create a PCF from scratch by reusing the code from the index.ts file, follow all the steps below.

I assume that you have the Microsoft Power Platform CLI installed in your machine. If not, please download it and install it before continue. Same for node.js

Just in case here are the links:

Power Platform CLI: https://aka.ms/PowerAppsCLI

Node.js: https://nodejs.org/es/download/


After you have installed Power Platform CLI and Node.js, open the Developer Command Prompt for Visual Studio. This is the tool which you will be using for most of the operations. First step is to create a folder where all component like files and folders will be present. 
1)	Manually create a folder
2)	Using the developer command prompt for visual studio apply the command: cd "folder path"
  
  The above command will take you to the folder created in the step 1. Next step is to create a new component project by passing the basic parameters using the below command. Component type parameter can be either field or dataset. Out of the two only field is available for canvas apps while both are available for model-driven apps.
  
3)	pac pcf init --namespace "specify your namespace here" --name "name of the component" --template "component type"
  
  Example: pac pcf init --namespace PCFFileUploader --name PCFFileUploader --template field
  
  Once the project is created, you have to bring in all the required project dependencies. For which you have to run the command as below
  
4)	npm install
  
  Then in order to complete the configuration of the PCF run the following commands
  
5)	Run: npm i dropzone@5.7.2
6)	Run: npm i toastr
7)	Run: npm i --save-dev @types/dropzone
8)	Run: npm i --save-dev @types/toastr
9)	npm install --save @types/xrm
  
  Next step is where you will define how the PCF will work and what functionality it will achieve. So to do that you have to open the folder in any IDE/text editor of your choice. I personally prefer VS Code.
  
10)	Copy the code from the index.ts file to your index.ts file
11)	Copy the missing resources and feature-usafe from the ControlManifest xml file
12)	Copy the css folder and img folder
13)	Check if the errors in the index.ts are gone, if the only error is the contactEntity from parameter, then run the following command: npm run build, after run this command, that error will disappear.
  
14)	Run: npm run build

Assuming that until here you are good and without errors, then the next step is to package the PCF in order to use it in your Dataverse environment. For that first create a new folder inside the component folder with a name of your choice. for example if the folder of step 1 was called PCFFileUploader, then create a folder called PCFFileUploaderSolution, the path would be ("root path"\PCFFileUploader\PCFFileUploader\PCFFileUploaderSolution)
  
Once the folder is created then using the command line, navigate to the newly created folder. From there only we will be running all our future commands. For example: cd "root path"\PCFFileUploader\PCFFileUploader\PCFFileUploaderSolution
  
Now you have to create a new solutions project using the below command. If you prefer to use any of the existing publisher, please give that name or else, you can provide the details of your choice but it should be unique to the environment.
  
15)	Run: pac solution init --publisher-name "publisher name" --publisher-prefix "prefix"

Next step is to add reference to the components which you created to the new solution project. You can use the below command for that. Make sure the path you are providing should be the folder where the component is located to be specific where your project file is located.
  
16)	Run: pac solution add-reference --path "root path"\PCFFileUploader

Next step is to create your solution file. For that run the command mentioned below. Only for the first time you have to use the restore parameter. After the first time, you can use msbuild /t:build and it should work.
  
17)	Run: msbuild /t:build /restore

Once the above command is executed successfully, you can find the solution filed in \bin\debug or \bin\release folder. Next step is to manually import the solution to your Dataverse environment through the portal just like you import any other solution.
  
18)	Upload the solution and publish it

Until here the PCF is already published in the environment. Now you have to create a Cloud flow because the PCF in the end is making a http request to a cloud flow.
Create a Cloud flow with an Http request trigger.
 
 ![image](https://user-images.githubusercontent.com/5630463/168445548-264fbeb6-f2d5-4475-b558-c4890563f06d.png)
 
And for the schema of the trigger you could use the following:
{
    "type": "object",
    "properties": {
        "filename": {
            "type": "string"
        },
        "filesize": {
            "type": "string"
        },
        "mimetype": {
            "type": "string"
        },
        "documentbody": {
            "type": "string"
        },
        "contactId": {
            "type": "string"
        },
        "userName": {
            "type": "string"
        },
        "noteText": {
            "type": "string"
        },
        "recordId": {
            "type": "string"
        },
        "contactName": {
            "type": "string"
        }
    }
}

Keep in mind that his payload is sent from the PCF, so if you want to add or remove parameters in this schema you have to change the code in the PCF to send the right payload to the cloud flow.
The function in the PCF which is in charge to send the http request is onButtonClick:

![image](https://user-images.githubusercontent.com/5630463/168445565-922cfa2e-7f93-46ce-9475-836df5dee18b.png)

Then you have to complete the cloud flow in order to create the file in SharePoint, for example:

 ![image](https://user-images.githubusercontent.com/5630463/168445577-14e930e8-5b0c-4abf-a893-aa95e25d8798.png)

As you can see, first I create the file then I update it with some properties, and again keep in mind that first you will have to create the columns in the SharePoint library and then you can update those columns in the cloud flow with the “Update file properties” action.

![image](https://user-images.githubusercontent.com/5630463/168445584-1f9a6cc1-7552-4d08-b6fb-0863e70981d9.png)

19)	Finally create a environment variable with the URL of the cloud flow, for example:

 ![image](https://user-images.githubusercontent.com/5630463/168445589-90433435-39da-4747-b636-b4283b390f36.png)

20)	And now you have all the components to allow the PCF upload a document to SharePoint. So the last step is go to the desired entity form and configure a text field like this:
 ![image](https://user-images.githubusercontent.com/5630463/168445517-554f04c1-09aa-4001-8e30-40ed215a259e.png)

There are three input parameters:
1)	A a lookup field. In my case I used the contact lookup field because I needed the GUID of the contact to create the right path of the file in the cloud flow in order to create it in SharePoint. If you don’t want a lookup field as input parameter, change the code of the PCF. 

3)	SupportedFiles, you could use this: 
application/pdf,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.ms-powerpoint,application/vnd.openxmlformats-officedocument.presentationml.presentation,image/png

If you want more type of files, please add it with a comma.

3)	EnvironmentVariableFlow, which should be the schema name of that variable.

21)	Save and publish the form

This is the PCF working:
![image](https://user-images.githubusercontent.com/5630463/168673152-0a7e2023-75ed-42ed-b58f-2db9a4ab72c5.png)


