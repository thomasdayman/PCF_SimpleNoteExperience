import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as $ from "jquery";

'use strict';
// Define const here

class NoteAttachment {
	fileName: string;
	mimeType: string;
	fileContent: string;
	fileSize: number;
}

let popUpContent: HTMLDivElement;

export class SimpleNoteExperience implements ComponentFramework.StandardControl<IInputs, IOutputs> {


	// Value of the field is stored and used inside the control 
	private _value: string | null;

	// PCF framework context, "Input Properties" containing the parameters, control metadata and interface functions.
	private _context: ComponentFramework.Context<IInputs>;

	// PCF framework delegate which will be assigned to this object which would be called whenever any update happens. 
	private _notifyOutputChanged: () => void;

	// Control's container
	private controlContainer: HTMLDivElement;

	private noteLabelElement: HTMLLabelElement;

	private _popUpService: ComponentFramework.FactoryApi.Popup.PopupService;
	private popUpButton: HTMLButtonElement;

	private entityName: string;
	private entityId: string;

	private loadingScreen: HTMLDivElement;
	private loadingSpinner: HTMLDivElement;

	private alertDialogDiv: HTMLDivElement;
	private successAlertDialog: HTMLDivElement;
	private failureAlertDialog: HTMLDivElement;

	private captureAudioButton: HTMLButtonElement;
	private captureImageButton: HTMLButtonElement;
	private captureVideoButton: HTMLButtonElement;
	private selectFileButton: HTMLButtonElement;
	private popUpContentButtonDiv: HTMLDivElement;
	private popUpNoteContentButtonDiv: HTMLDivElement;

	private addNoteButton: HTMLButtonElement;
	private saveNoteButton: HTMLButtonElement;
	private cancelNoteButton: HTMLButtonElement;

	private AddNoteSubjectBox: HTMLInputElement;
	private AddNoteTextBox: HTMLTextAreaElement;
	private attachFileButton: HTMLButtonElement;

	private NoteAttachmentObject = {} as NoteAttachment;

	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		interface PopupDev extends ComponentFramework.FactoryApi.Popup.Popup {
			popupStyle: object;
		}
		this._context = context;
		this._notifyOutputChanged = notifyOutputChanged;

		this.entityName = (context.mode as any).contextInfo.entityTypeName;
		this.entityId = (context.mode as any).contextInfo.entityId;

		this.controlContainer = document.createElement("div");

		this.loadingScreen = document.createElement("div");
		this.loadingScreen.id = "loader";

		this.loadingSpinner = document.createElement("div");
		this.loadingSpinner.id = "loadingSpinner";

		this.successAlertDialog = document.createElement("div");
		this.successAlertDialog.id = "success";

		this.failureAlertDialog = document.createElement("div");
		this.failureAlertDialog.id = "failure";

		this.alertDialogDiv = document.createElement("div");
		this.alertDialogDiv.id = "mainAlertDiv"
		this.alertDialogDiv.appendChild(this.successAlertDialog);
		this.alertDialogDiv.appendChild(this.failureAlertDialog);

		this.popUpButton = document.createElement("button");
		this.popUpButton.innerHTML = "Capture Media";
		this.popUpButton.addEventListener("click", this.onPopUpButtonClick.bind(this));

		this.captureAudioButton = document.createElement("button");
		this.captureAudioButton.innerHTML = "Capture Audio";
		this.captureAudioButton.addEventListener("click", this.onCaptureAudioClick.bind(this));

		this.captureImageButton = document.createElement("button");
		this.captureImageButton.innerHTML = "Capture Image";
		this.captureImageButton.addEventListener("click", this.onCaptureImageClick.bind(this));

		this.captureVideoButton = document.createElement("button");
		this.captureVideoButton.innerHTML = "Capture Video";
		this.captureVideoButton.addEventListener("click", this.onCaptureVideoClick.bind(this));

		this.selectFileButton = document.createElement("button");
		this.selectFileButton.innerHTML = "Add File";
		this.selectFileButton.addEventListener("click", this.onSelectFileClick.bind(this));

		this.addNoteButton = document.createElement("button");
		this.addNoteButton.innerHTML = "Add Note";
		this.addNoteButton.addEventListener("click", this.onAddNoteClick.bind(this));

		this.popUpContentButtonDiv = document.createElement('div');
		this.popUpContentButtonDiv.id = "mainButtonDiv";
		this.popUpContentButtonDiv.appendChild(this.captureAudioButton);
		this.popUpContentButtonDiv.appendChild(this.captureImageButton);
		this.popUpContentButtonDiv.appendChild(this.captureVideoButton);
		this.popUpContentButtonDiv.appendChild(this.selectFileButton);
		this.popUpContentButtonDiv.appendChild(this.addNoteButton);

		this.AddNoteSubjectBox = document.createElement("input");
		this.AddNoteSubjectBox.placeholder = "Enter Subject...";

		this.AddNoteTextBox = document.createElement("textarea");
		this.AddNoteTextBox.placeholder = "Enter Text...";

		this.noteLabelElement = document.createElement("label");
		this.noteLabelElement.innerHTML = "Attachment...";

		this.attachFileButton = document.createElement("button");
		this.attachFileButton.innerHTML = "Attach File...";
		this.attachFileButton.addEventListener("click", this.onattachFileButtonClick.bind(this));

		this.saveNoteButton = document.createElement("button");
		this.saveNoteButton.innerHTML = "Create Note";
		this.saveNoteButton.addEventListener("click", this.onSaveAddNoteClick.bind(this));

		this.cancelNoteButton = document.createElement("button");
		this.cancelNoteButton.innerHTML = "Cancel";
		this.cancelNoteButton.style.backgroundColor = "#C21807"
		this.cancelNoteButton.addEventListener("click", this.onCancelNoteClick.bind(this));

		this.popUpNoteContentButtonDiv = document.createElement('div');
		this.popUpNoteContentButtonDiv.id = "noteButtonDiv"
		this.popUpNoteContentButtonDiv.appendChild(this.AddNoteSubjectBox);
		this.popUpNoteContentButtonDiv.appendChild(this.AddNoteTextBox);
		this.popUpNoteContentButtonDiv.appendChild(this.noteLabelElement);
		this.popUpNoteContentButtonDiv.appendChild(this.attachFileButton);
		this.popUpNoteContentButtonDiv.appendChild(this.saveNoteButton);
		this.popUpNoteContentButtonDiv.appendChild(this.cancelNoteButton);

		popUpContent = document.createElement('div');
		popUpContent.id = "popUp";
		popUpContent.style.backgroundColor = "white";
		popUpContent.appendChild(this.alertDialogDiv);
		popUpContent.appendChild(this.popUpContentButtonDiv);

		let popUpOptions: PopupDev = {
			closeOnOutsideClick: true,
			content: popUpContent,
			name: 'mainPopup',
			type: 1,
			popupStyle: {}
		};

		this.controlContainer.appendChild(this.popUpButton);
		container.appendChild(this.controlContainer);
		this._popUpService = context.factory.getPopupService();
		this._popUpService.createPopup(popUpOptions);
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		this._context = context;
	}

	public showLoading(): void {
		this._popUpService.updatePopup("mainPopup", {
			closeOnOutsideClick: false,
			content: popUpContent,
			name: 'mainPopup',
			type: 1,
		});
		popUpContent.prepend(this.loadingScreen);
		popUpContent.prepend(this.loadingSpinner);
	}

	public stopLoading(): void {
		this._popUpService.updatePopup("mainPopup", {
			closeOnOutsideClick: true,
			content: popUpContent,
			name: 'mainPopup',
			type: 1,
		});
		this.loadingScreen.remove();
		this.loadingSpinner.remove();
	}

	public showMainPopUp(): void {
		popUpContent.removeChild(this.popUpNoteContentButtonDiv);
		popUpContent.appendChild(this.popUpContentButtonDiv);
	}

	public showNotePopUp(): void {
		popUpContent.removeChild(this.popUpContentButtonDiv);
		popUpContent.appendChild(this.popUpNoteContentButtonDiv);
	}

	private onPopUpButtonClick(event: Event): void {
		this._popUpService.openPopup('mainPopup');
		this.showMainPopUp();
	}

	private onSelectFileClick(event: Event): void {
		this._context.device.pickFile().then(this.quickUploadPickFile.bind(this), this.showFailureAlert.bind(this));
	}

	private onattachFileButtonClick(event: Event): void {
		this._context.device.pickFile().then(this.storeAttachedFile.bind(this), this.showFailureAlert.bind(this));
	}

	private onCaptureAudioClick(event: Event): void {
		try {
			this._context.device.captureAudio().then(this.quickUploadCapture.bind(this), this.showFailureAlert.bind(this));
		}
		catch (err) {
			this.showFailureAlert(err);
		}
	}

	private onCaptureVideoClick(event: Event): void {
		try {
			this._context.device.captureVideo().then(this.quickUploadCapture.bind(this), this.showFailureAlert.bind(this));
		}
		catch (err) {
			this.showFailureAlert(err);
		}
	}

	private onCaptureImageClick(event: Event): void {
		try {
			this._context.device.captureImage().then(this.quickUploadCapture.bind(this), this.showFailureAlert.bind(this));
		}
		catch (err) {
			this.showFailureAlert(err);
		}
	}

	private onAddNoteClick(event: Event): void {
		this.showNotePopUp();
	}

	private onCancelNoteClick(event: Event): void {
		this.showMainPopUp();
		this.clearNoteAttachment();
	}

	private clearNoteAttachment(): void {
		this.NoteAttachmentObject = new NoteAttachment()
		this.noteLabelElement.innerHTML = "Attachment...";
		this.AddNoteTextBox.value = "";
		this.AddNoteTextBox.placeholder = "Enter Text...";
		this.AddNoteSubjectBox.value = "";
		this.AddNoteSubjectBox.placeholder = "Enter Subject...";
	}

	private onSaveAddNoteClick(event: Event): void {
		this.showLoading();
		if (this.NoteAttachmentObject != null) {
			this.uploadNote(true);
		}
		else {
			this.uploadNote(false);
		}
	}

	private storeAttachedFile(files: ComponentFramework.FileObject[]): void {
		if (files.length > 0) {
			let file: ComponentFramework.FileObject = files[0];

			try {
				if (file && file.fileName) {
					this.setNoteLabel(file.fileName);

					let NoteAttachmentObject: NoteAttachment = {
						fileName: file.fileName,
						mimeType: file.mimeType,
						fileContent: file.fileContent,
						fileSize: file.fileSize,
					};
					this.NoteAttachmentObject = NoteAttachmentObject;
				}
				else {
					this.stopLoading();
					this.showFailureAlert(Error("Attachment was not found..."));
				}
			}
			catch (err) {
				this.showFailureAlert(err);
			}
		}
	}

	private quickUploadPickFile(files: ComponentFramework.FileObject[]): void {
		this.showLoading();
		if (files.length > 0) {
			let file: ComponentFramework.FileObject = files[0];
			try {
				if (file && file.fileName) {
					this.createNote(file.fileName, "", true, this.entityName, this.entityId, file.fileName, file.mimeType, file.fileContent, file.fileSize, true);
				}
				else {
					this.stopLoading();
					this.showFailureAlert(Error("Media was not found..."));
				}
			}
			catch (err) {
				this.stopLoading();
				this.showFailureAlert(err);
			}
		}
	}

	private quickUploadCapture(file: ComponentFramework.FileObject): void {
		this.showLoading();
		try {
			if (file && file.fileName) {
				this.createNote(file.fileName, "", true, this.entityName, this.entityId, file.fileName, file.mimeType, file.fileContent, file.fileSize, true);
			}
			else {
				this.stopLoading();
				this.showFailureAlert(Error("Media was not found..."));
			}
		}
		catch (err) {
			this.stopLoading();
			this.showFailureAlert(err);
		}
	}

	private uploadNote(isDocument: boolean): void {
		let notetext = this.AddNoteTextBox.value
		let subject = this.AddNoteSubjectBox.value
		this.createNote(subject, notetext, isDocument, this.entityName, this.entityId, this.NoteAttachmentObject.fileName, this.NoteAttachmentObject.mimeType, this.NoteAttachmentObject.fileContent, this.NoteAttachmentObject.fileSize, false);
	}

	private showFailureAlert(error: Error): void {
		$("#failure").fadeIn(300).delay(2000).fadeOut(400);
		this.failureAlertDialog.innerHTML = error.message;
	}

	private showSuccessAlert(success: any): void {
		$("#success").fadeIn(300).delay(2000).fadeOut(400);
		this.successAlertDialog.innerHTML = success;
	}

	private setNoteLabel(fileName: string): void {
		this.noteLabelElement.innerHTML = fileName;
	}

	private createNote(subject: string, notetext: string, isDocument: boolean, entityName: string | null, entityId: string | null, fileName: string | null, mimeType: string | null, fileContent: string | null, fileSize: Number, isMainMenu: boolean): void {
		let entity: any = {};
		entity.subject = subject;
		entity.notetext = notetext;
		if (fileName != null) {
			entity.filename = fileName;
			entity.isdocument = isDocument;
			entity.documentbody = fileContent;
			entity.mimetype = mimeType;
		}
		entity["objectid_" + entityName + "@odata.bind"] = "/" + entityName + "s(" + entityId + ")";

		this._context.webAPI.createRecord("annotation", entity).then(
			(result) => {
				var newEntityId = result.id;
				if (isMainMenu == false) {
					this.clearNoteAttachment();
					this.showMainPopUp();
				}
				this.stopLoading();
				this.showSuccessAlert("Note was successfully created!");
			},
			(error) => {
				this.stopLoading();
				this.showFailureAlert(error);
			}
		);
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
		 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
	}

}