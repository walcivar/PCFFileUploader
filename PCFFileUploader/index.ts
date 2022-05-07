import {IInputs, IOutputs} from "./generated/ManifestTypes";

import * as Dropzone from "dropzone";
import * as toastr from "toastr";

//Supported files
const supportedfiles = "application/pdf,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet," +
    "application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.ms-powerpoint," +
    "application/vnd.openxmlformats-officedocument.presentationml.presentation";

class EntityReference
{
    id: string;
    typeName: string;
    constructor(typeName: string, id: string) 
    {
        this.id = id;
        this.typeName = typeName;
    }
}

class AttachedFile implements ComponentFramework.FileObject
{
    fileId: string;
    fileContent: string;
    fileSize: number;
    fileName: string;
    mimeType: string;
    constructor(fileId: string, fileName: string, mimeType: string, fileContent: string, fileSize: number)
    {
        this.fileId = fileId
        this.fileName = fileName;
        this.mimeType = mimeType;
        this.fileContent = fileContent;
        this.fileSize = fileSize;
    }
}

export class PCFFileUploader implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private entityReference: EntityReference;

    private _context: ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;

    private _divDropZone: HTMLDivElement;
    private _formDropZone: HTMLFormElement;
    private _imgUpload: HTMLImageElement;

    private _brElement: HTMLBRElement;

    private _divFile: HTMLDivElement;
    private _labelFile: HTMLLabelElement;

    private _reason: HTMLSelectElement;
    private _doctype: HTMLSelectElement;
    private _textAreaNote: HTMLTextAreaElement;
    private _buttonFlow: HTMLButtonElement;
    private _buttonClose: HTMLButtonElement;

    private _listReason: ComponentFramework.WebApi.Entity[];
    private _listDocType: ComponentFramework.WebApi.Entity[];

    private _attachedFile: any = {}
    private _urlSharepoint: string;

    private _userName: string;
    private _contactGuid: string;
    private _contactName: string;

    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        this._context = context;

        this.entityReference = new EntityReference(
            (<any>context).page.entityTypeName,
            (<any>context).page.entityId
        )

        this._container = document.createElement("div");
        this._brElement = document.createElement("br");

        this._imgUpload = document.createElement("img");
        this._imgUpload.setAttribute("class", "uploadImgHowerOut");

        this._context.resources.getResource("newUploadIcon.png", data => {
            let uploadImage = this.generateSrcUrl("image", "png", data);
            this._imgUpload.src = uploadImage;
        }, () => {
            console.log("upload image loading error");
        });

        this._imgUpload.addEventListener("mouseover", this.uploadImgHower.bind(this));
        this._imgUpload.addEventListener("mouseout", this.uploadImgHowerOut.bind(this));
        this._imgUpload.addEventListener("click", this.uploadImgonClick.bind(this));

        this._divDropZone = document.createElement("div");
        this._divDropZone.id = "dropzone";

        this._formDropZone = document.createElement("form");
        this._formDropZone.id = "upload_dropzone";
        this._formDropZone.setAttribute("method", "post");

        this._formDropZone.appendChild(this._imgUpload);
        this._formDropZone.setAttribute("class", "dropzone needsclick");
        this._divDropZone.appendChild(this._formDropZone);

        this._divFile = document.createElement("div");

        this._container.appendChild(this._divDropZone);
        this._container.appendChild(this._divFile);
        container.appendChild(this._container);

        this.onload();

        toastr.options.closeButton = true;
        toastr.options.progressBar = true;
        toastr.options.positionClass = "toast-bottom-left";

        const fetchSharepointParam = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
            "<entity name='environmentvariablevalue'>" +
                "<attribute name='environmentvariablevalueid' />" +
                "<attribute name='value' />" +
                "<link-entity name='environmentvariabledefinition' from='environmentvariabledefinitionid' to='environmentvariabledefinitionid' link-type='inner' alias='ag'>" +
                    "<filter type='and'>" +
                        "<condition attribute='schemaname' operator='eq' value='p3i_sharepointflowurl' />" +
                    "</filter>" +
                "</link-entity>" +
            "</entity>" +
        "</fetch>";

        let thisRefTemp = this;

        this._context.webAPI.retrieveMultipleRecords("environmentvariablevalue", "?fetchXml=" + fetchSharepointParam).then
            (
            function (params: ComponentFramework.WebApi.RetrieveMultipleResponse) {

                thisRefTemp._urlSharepoint = params.entities[0].value;
            },
            function (error: any) {

                console.log("Error recuperando la variable de entorno");
            }
        );

        //Get all reasons
        const fetchMotivos = "<fetch>" +
            "<entity name='p3i_reason' >" +
            "   <attribute name='p3i_name' />" +
            "   <attribute name='p3i_code' />" +
            "</entity>" +
            "</fetch>";
        
        this._context.webAPI.retrieveMultipleRecords("p3i_reason", "?fetchXml=" + fetchMotivos).then
        (
        function (params: ComponentFramework.WebApi.RetrieveMultipleResponse) {

            thisRefTemp._listReason = params.entities;       
        },
        function (error: any) {

            console.log("Error recuperando los Motivos");
        });

        //Get all document types
        const fetchTipoDocumentos = "<fetch>" +
            "<entity name='p3i_documenttype' >" +
            "   <attribute name='p3i_name' />" +
            "   <attribute name='p3i_code' />" +
            "</entity>" +
            "</fetch>";

        this._context.webAPI.retrieveMultipleRecords("p3i_documenttype", "?fetchXml=" + fetchTipoDocumentos).then
        (
        function (params: ComponentFramework.WebApi.RetrieveMultipleResponse) {

            thisRefTemp._listDocType = params.entities;                       
        },
        function (error: any) {

            console.log("Error recuperando los Tipos de documentos");
        });

        this._userName = Xrm.Utility.getGlobalContext().userSettings.userName;
        const contactRecord = this._context.parameters.contactEntity.raw ? this._context.parameters.contactEntity.raw : undefined;
        if(contactRecord)
        {
            this._contactGuid = contactRecord[0].id;
            this._contactName = contactRecord[0].name!;
        } 
        //this._contactGuid = Xrm.Page.getAttribute("p3i_contacto").getValue()[0].id;
               
        console.log("Contact GUID method 2 : "+ this._contactGuid);
    }

    private onload() 
    {
        let thisRef = this;
        new Dropzone(this._formDropZone, 
        {
            acceptedFiles: supportedfiles,
            url: "/",
            parallelUploads: 2,
            maxFilesize: 3,
            filesizeBase: 1000,
            dictDefaultMessage: "Drag the files you want to upload here",
            addedfile: function (file: any) {
                if (supportedfiles.indexOf(file.type) == -1 ) { return; }
                const fileSize = file.upload.total;
                const fileName = file.upload.filename;
                const mimeType = file.type;
                const reader = new FileReader();
                reader.onload = function (event: any) {
                    const fileContent = event.target.result;
                    const attachedFile = new AttachedFile("", fileName, mimeType, fileContent, fileSize);
                    thisRef.addAttachments(attachedFile);
                };
                reader.readAsDataURL(file);
            },
        });
    }

    private addAttachments(file: AttachedFile): void 
    {
        //Remove the div with the attributes if existing
        if (this._divFile.childElementCount > 0)
        {
            this._divFile.children[0].remove();
        }

        let attachedFile: any = {}

        const index = file.fileContent.indexOf(";base64,");
        const fileContent = file.fileContent.substring(index + 8);

        if (fileContent != null || "" || undefined && file.fileName != null || "" || undefined) 
        {
            attachedFile["documentbody"] = fileContent;
            attachedFile["fileName"] = file.fileName;
            attachedFile["filesize"] = file.fileSize;
            attachedFile["mimetype"] = file.mimeType;
        }
        
        let thisRef = this;

        thisRef.addFileControl(attachedFile);
        this._attachedFile = attachedFile;
        toastr.success("The file has been selected correctly.<br/>Select the corresponding attributes before uploading.");
    }

    private addFileControl(FileToUpload: AttachedFile) 
    {
        this._labelFile = document.createElement("label");
        this._labelFile.className = "text-font";
        this._labelFile.innerHTML = FileToUpload.fileName;

        let dropdownmotivo = document.createElement("select");
        dropdownmotivo.id = "p3i_reasonid";
        this._reason = dropdownmotivo;

        let dropdowntipodocumento = document.createElement("select");
        dropdowntipodocumento.id = "p3i_documenttypeid";
        this._doctype = dropdowntipodocumento;

        this._listReason.forEach(function (motivo)
        {
            let option = document.createElement("option");
            option.value = motivo.p3i_code;
            option.text = motivo.p3i_name;
            dropdownmotivo.appendChild(option);
        });

        this._listDocType.forEach(function (tipoDoc) 
        {
            let option = document.createElement("option");
            option.value = tipoDoc.p3i_code;
            option.text = tipoDoc.p3i_name;
            dropdowntipodocumento.appendChild(option);
        });
       
        this._textAreaNote = document.createElement("textarea");
        this._textAreaNote.rows = 4;
        this._textAreaNote.maxLength = 500;
        this._textAreaNote.className = "noteText";
        this._textAreaNote.id = "notaid";

        this._buttonFlow = document.createElement("button");
        this._buttonFlow.innerHTML = "Send to SP";
        this._buttonFlow.className = "boton";
        this._buttonFlow.addEventListener("click", this.onButtonClick.bind(this));

        this._buttonClose = document.createElement("button");
        this._buttonClose.innerHTML = "Cancel";
        this._buttonClose.addEventListener("click", this.onButtonClickClose.bind(this));
        this._buttonClose.className = "boton";


        let _divFileContainer = document.createElement("div");
        _divFileContainer.id = "divFileContainer_" + FileToUpload.fileId;
        
        this._divFile.appendChild(_divFileContainer);        
        
        _divFileContainer.innerHTML = "<br/><table style='width:100%;'>"+
            "<tr>" +
            "<td>Document Type</td>" +
            "</tr>" +
            "<tr>" +
            "<td id='docTypeCell'></td>" +
            "</tr>" +
            "<tr>" +
            "<td>Reason</td>" +
            "</tr>" +
            "<tr>" +
            "<td id='reasonCell' style='width:100%'></td>" +
            "</tr>" +
            "<tr>" +
            "<td>Notes</td>" +
            "</tr>" +
            "<tr>" +
            "<td id='notesCell' style='width:100%'></td>" +
            "</tr>" +
            "</table>" +
            "<br/>";
        
        $("#docTypeCell")[0].appendChild(this._doctype);
        $("#reasonCell")[0].appendChild(this._reason);
        $("#notesCell")[0].appendChild(this._textAreaNote);

        $("#divFileContainer_" + FileToUpload.fileId)[0].appendChild(this._buttonFlow);
        $("#divFileContainer_" + FileToUpload.fileId)[0].appendChild(this._buttonClose);
        $("#divFileContainer_" + FileToUpload.fileId)[0].appendChild(this._brElement.cloneNode());
        $("#divFileContainer_" + FileToUpload.fileId)[0].appendChild(this._brElement.cloneNode());
        $("#divFileContainer_" + FileToUpload.fileId)[0].appendChild(this._labelFile);
    }

    private onButtonClickClose(event: Event): void 
    {
        this._divFile.children[0].remove();
    }

    private onButtonClick(event: Event): void
    {
        const reasonText = $("#p3i_reasonid option:selected").text();
        const reasonId = $("#p3i_reasonid option:selected").val();
        const doctypeText = $("#p3i_documenttypeid option:selected").text();
        const doctypeId = $("#p3i_documenttypeid option:selected").val();
        const noteText = $("#notaid").val() != null ? $("#notaid").val() : "";

        const req = new XMLHttpRequest();
        const url = this._urlSharepoint;
        req.open("POST", url, true);
        req.setRequestHeader('Content-Type', 'application/json');
        req.send(JSON.stringify({
            "filename": "" + this._attachedFile["fileName"] + "",
            "filesize": "" + this._attachedFile["filesize"] + "",
            "mimetype": "" + this._attachedFile["mimetype"] + "",
            "documentbody": "" + this._attachedFile["documentbody"] + "",
            "reasonText": "" + reasonText + "",
            "reasonId": "" + reasonId + "",
            "doctypeText": "" + doctypeText + "",
            "doctypeId": "" + doctypeId + "",
            "recordId": "" + this.entityReference.id + "",
            "contactId": ""+ this._contactGuid +"",
            "contactName": ""+ this._contactName +"",
            "userName": "" + this._userName +"",
            "noteText": "" + noteText + ""
        }));
        toastr.success("File sent to \"Power Automate\" to be processed.<br/>Soon will be in Sharepoint");
        //Remove the div with the attributes
        this._divFile.children[0].remove();
    }

    private generateSrcUrl(datatype: string, fileType: string, fileContent: string): string {
        return "data:" + datatype + "/" + fileType + ";base64, " + fileContent;
    }

    private uploadImgHower() {
        this._imgUpload.setAttribute("class", "uploadImgHower");
    }

    private uploadImgHowerOut() {
        this._imgUpload.setAttribute("class", "uploadImgHowerOut");
    }

    private uploadImgonClick() {
        this._formDropZone.click();
    }

    private CollectionNameFromLogicalName(entityLogicalName: string): string 
    {
        if (entityLogicalName[entityLogicalName.length - 1] != 's') {
            return `${entityLogicalName}s`;
        }
        else 
        {
            return `${entityLogicalName}es`;
        }
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        // Add code to update control view
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }
}
