import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface AbnRecord {
	EntityName: string;
	AbnCode: string
}

export class ABNChecker implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private notifyOutputChanged: () => void;
	private container: HTMLDivElement;
	private baseUrl: string;
	private context: ComponentFramework.Context<IInputs>;
	private inputControl: HTMLInputElement;
	private currentFocus: any;
	private records: AbnRecord[] = [];
	private abnFieldValue: string;
	private guid: string;
	private isAbnNumber: boolean;

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
		// Add control initialization code
		this.notifyOutputChanged = notifyOutputChanged;

		this.container = container;
		this.baseUrl = 'https://licensing.maptaskr.com'; //'http://localhost:55891';
		this.guid = 'd998c3b3-fa2b-47a4-865f-dc3b9bc00a24';
		this.context = context;
		let inputContainer = document.createElement("div");
		inputContainer.className = "abn-autocomplete";

		this.inputControl = document.createElement("input");
		this.inputControl.placeholder = '---';
		this.inputControl.addEventListener("input", this.onChange.bind(this));
		this.inputControl.addEventListener("keydown", this.onKeyDown.bind(this));

		document.addEventListener("click", this.onDocumentClick.bind(this))

		inputContainer.appendChild(this.inputControl);
		container.appendChild(inputContainer);
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view
		this.context = context;
		this.abnFieldValue = this.context.parameters.AbnField.raw || "";
		this.inputControl.value = this.abnFieldValue;
		if (this.abnFieldValue.length > 2) this.populateRecords();
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {
			AbnField: this.abnFieldValue
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
		this.inputControl.removeEventListener("input", this.onChange);
		this.inputControl.removeEventListener("keydown", this.onKeyDown);
		document.removeEventListener("click", this.onDocumentClick);
	}

	/**
	 * Populate entities
	 */
	private async populateRecords() {
		this.records = [];

		/*close any already open lists of autocompleted values*/
		this.closeAllLists();

		if (this.abnFieldValue.length <= 2) return;

		this.isAbnNumber = this.validateAbnNumber();
		let jsonNames = await this.getNames();
		let nameRecords = jsonNames.replace('callback(', '').slice(0, -1);
		if (!nameRecords) return;

		switch (true) {
			case this.isAbnNumber == false:
				var result = JSON.parse(nameRecords);
				var matchedNames = result ? result.Names : undefined;
				this.pushNames(matchedNames);
				break;
			default:
				var result = JSON.parse(nameRecords);
				this.records.push(<AbnRecord>{
					EntityName: result.EntityName,
					AbnCode: result.Abn
				});
				break;
		}

		if (this.records && this.abnFieldValue.length > 2) {
			this.populateAutoComplete();
		}
	}

	private validateAbnNumber() {
		return (Number.isInteger(parseInt(this.abnFieldValue)) && parseInt(this.abnFieldValue).toString().length == 11);
	}

	private pushNames(matchedNames: any) {
		for (var i = 0; i < matchedNames.length; i++) {
			if (matchedNames[i].Name) {
				this.records.push(<AbnRecord>{
					EntityName: matchedNames[i].Name,
					AbnCode: matchedNames[i].Abn
				});
			}
		}
	}

	/**
	 * Get entities 
	 */
	private async getNames(): Promise<string> {
		const requestUrl = `${this.baseUrl}/api/Abn/MatchingNames/eyJhbGciOi/${this.abnFieldValue}`;
		return this.request(requestUrl);
	}

	/**
	 * Common call back for api
	 * @param param 
	 */
	private request(requestUrl: string): Promise<string> {
		var that = this;

		return new Promise(function (resolve, reject) {

			var xhr = new XMLHttpRequest();

			xhr.addEventListener("readystatechange", function () {
				if (this.readyState === 4) {
					resolve(this.responseText);
				}
			});

			xhr.open("GET", requestUrl);

			xhr.send();
		});
	}

	/**
	 * Called when input change
	 * @param evt 
	 */
	private onChange(evt: any): void {
		this.abnFieldValue = (this.inputControl.value as any) as string;
		this.notifyOutputChanged();
	}

	/**
	 * Populate suggestion
	 */
	private populateAutoComplete() {
		const that = this;
		var a, b, i, val = (that.inputControl.value as any) as string;

		if (!val) {
			return;
		}
		that.currentFocus = -1;
		/*create a DIV element that will contain the items (values):*/
		a = document.createElement("DIV");
		a.setAttribute("id", that.inputControl.id + "autocomplete-list");
		a.setAttribute("class", "abn-autocomplete-items");
		/*append the DIV element as a child of the autocomplete container:*/
		(that.inputControl as any).parentNode.appendChild(a);
		/*for each item in the array...*/
		for (i = 0; i < that.records.length; i++) {
			/*check if the item starts with the same letters as the text field value:*/
			const isMatched = this.isAbnNumber ? that.records[i].AbnCode.substr(0, val.length).toUpperCase() == val.toUpperCase()
				: that.records[i].EntityName.substr(0, val.length).toUpperCase() == val.toUpperCase();

			if (isMatched) {
				/*create a DIV element for each matching element:*/
				b = document.createElement("DIV");
				/*make the matching letters bold:*/
				b.innerHTML = that.records[i].AbnCode + ' ' + "<strong>" + that.records[i].EntityName.substr(0, val.length) + "</strong>";
				b.innerHTML += that.records[i].EntityName.substr(val.length);
				/*insert a input field that will hold the current array item's value:*/
				b.innerHTML += "<input type='hidden' value='" + that.records[i].AbnCode + "'>";
				/*execute a function when someone clicks on the item value (DIV element):*/
				b.addEventListener("click", function (e) {
					/*insert the value for the autocomplete text field:*/
					that.inputControl.value = this.getElementsByTagName("input")[0].value;
					that.abnFieldValue = that.inputControl.value;
					that.closeAllLists();
				});
				a.appendChild(b);
			}
		}
	}

	/**
	 * Called when key down on input 
	 * @param evt 
	 */
	private onKeyDown(evt: any): void {
		const that = this;
		var x = <any>document.getElementById(evt.id + "autocomplete-list");
		if (x) x = x.getElementsByTagName("div");
		if (evt.keyCode == 40) {
			/*If the arrow DOWN key is pressed,
			increase the currentFocus variable:*/
			that.currentFocus++;
			/*and and make the current item more visible:*/
			that.addActive(x);
		} else if (evt.keyCode == 38) { //up
			/*If the arrow UP key is pressed,
			decrease the currentFocus variable:*/
			that.currentFocus--;
			/*and and make the current item more visible:*/
			that.addActive(x);
		} else if (evt.keyCode == 13) {
			/*If the ENTER key is pressed, prevent the form from being submitted,*/
			evt.preventDefault();
			if (that.currentFocus > -1) {
				/*and simulate a click on the "active" item:*/
				if (x) x[that.currentFocus].click();
			}
		}
	}

	/**
	 * Called when document click
	 * @param evt 
	 */
	private onDocumentClick(evt: any): void {
		const that = this;
		that.closeAllLists(evt.target);
	}

	/**
	 * 
	 * @param x 
	 */
	private addActive(x: any) {
		/*a function to classify an item as "active":*/
		if (!x) return false;
		/*start by removing the "active" class on all items:*/
		this.removeActive(x);
		if (this.currentFocus >= x.length) this.currentFocus = 0;
		if (this.currentFocus < 0) this.currentFocus = (x.length - 1);
		/*add class "autocomplete-active":*/
		x[this.currentFocus].classList.add("abn-autocomplete-active");
	}

	/**
	 * 
	 * @param x 
	 */
	private removeActive(x: any) {
		/*a function to remove the "active" class from all autocomplete items:*/
		for (var i = 0; i < x.length; i++) {
			x[i].classList.remove("abn-autocomplete-active");
		}
	}

	/**
	 * 
	 * @param elmnt 
	 */
	private closeAllLists(elmnt?: any) {
		/*close all autocomplete lists in the document,
		except the one passed as an argument:*/
		var x = <any>document.getElementsByClassName("abn-autocomplete-items");
		for (var i = 0; i < x.length; i++) {
			if (elmnt != x[i] && elmnt != this.inputControl) {
				x[i].parentNode.removeChild(x[i]);
			}
		}
	}
}