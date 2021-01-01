import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class DataSourceAutoComplete implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private notifyOutputChanged: () => void;
	private container: HTMLDivElement;
	private baseUrl: string;
	private context: ComponentFramework.Context<IInputs>;
	private inputControl: HTMLInputElement;
	private currentFocus: any;
	private records: any[] = [];
	private datasources: any[] = [];
	private rlEntities: any[] = [];
	private value: string;
	private type: string;
	private entity: string;
	private formContext: any;
	private fieldLogicalName: string;

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
		this.baseUrl = (<any>context).page.getClientUrl();
		this.context = context;
		let inputContainer = document.createElement("div");
		inputContainer.className = "autocomplete";

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
		this.type = this.context.parameters.Type.raw || "1";
		this.value = this.context.parameters.Field.raw || "";
		this.inputControl.value = this.value;

		this.fieldLogicalName = this.context.parameters.Field.attributes ?
			this.context.parameters.Field.attributes.LogicalName : '';

		this.setFormContext();
		this.setParentEntity();

		switch (this.type) {
			case '1':
				if (this.records.length === 0) this.populateEntities();
				break;
			case '2':
				if (this.entity && this.entity.length > 0 && this.records.length === 0) this.populateAttributes(this.entity);
				break;
			case '3':
				if (this.entity && this.entity.length > 0 && this.records.length === 0) this.populateRelatedEntities(this.entity);
				break;
		}
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {
			Field: this.value
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
	 * Called when input change
	 * @param evt 
	 */
	private onChange(evt: any): void {

		if (this.records) {
			this.populateAutoComplete();
		}

		this.notifyOutputChanged();
	}

	/**
	 * Populate suggestion
	 */
	private populateAutoComplete() {
		const that = this;
		var a, b, i, val = (that.inputControl.value as any) as string;

		that.value = val;

		/*close any already open lists of autocompleted values*/
		that.closeAllLists();
		if (!val) {
			return;
		}
		that.currentFocus = -1;
		/*create a DIV element that will contain the items (values):*/
		a = document.createElement("DIV");
		a.setAttribute("id", that.inputControl.id + "autocomplete-list");
		a.setAttribute("class", "autocomplete-items");
		/*append the DIV element as a child of the autocomplete container:*/
		(that.inputControl as any).parentNode.appendChild(a);
		/*for each item in the array...*/
		for (i = 0; i < that.records.length; i++) {
			/*check if the item starts with the same letters as the text field value:*/
			if (that.records[i].substr(0, val.length).toUpperCase() == val.toUpperCase()) {
				/*create a DIV element for each matching element:*/
				b = document.createElement("DIV");
				/*make the matching letters bold:*/
				b.innerHTML = "<strong>" + that.records[i].substr(0, val.length) + "</strong>";
				b.innerHTML += that.records[i].substr(val.length);
				/*insert a input field that will hold the current array item's value:*/
				b.innerHTML += "<input type='hidden' value='" + that.records[i] + "'>";
				/*execute a function when someone clicks on the item value (DIV element):*/
				b.addEventListener("click", function (e) {
					/*insert the value for the autocomplete text field:*/
					that.inputControl.value = this.getElementsByTagName("input")[0].value;
					that.value = that.inputControl.value;
					that.setFormContext();
					if (that.fieldLogicalName) that.formContext.getAttribute(that.fieldLogicalName).setValue(that.value);

					if (that.type == '3') that.onRelatedEntitySelected();
					/*close the list of autocompleted values,
					(or any other open lists of autocompleted values:*/
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
		x[this.currentFocus].classList.add("autocomplete-active");
	}

	/**
	 * 
	 * @param x 
	 */
	private removeActive(x: any) {
		/*a function to remove the "active" class from all autocomplete items:*/
		for (var i = 0; i < x.length; i++) {
			x[i].classList.remove("autocomplete-active");
		}
	}

	/**
	 * 
	 * @param elmnt 
	 */
	private closeAllLists(elmnt?: any) {
		/*close all autocomplete lists in the document,
		except the one passed as an argument:*/
		var x = <any>document.getElementsByClassName("autocomplete-items");
		for (var i = 0; i < x.length; i++) {
			if (elmnt != x[i] && elmnt != this.inputControl) {
				x[i].parentNode.removeChild(x[i]);
			}
		}
	}

	/**
	 * Populate entities
	 */
	private async populateEntities() {
		this.records = [];
		var a = await this.getEntities();
		var result = JSON.parse(a);

		for (var i = 0; i < result.value.length; i++) {
			if (result.value[i].LogicalName !== null) {
				this.records.push(result.value[i].LogicalName);
			}
		}
	}

	/**
	 * 
	 */

	onRelatedEntitySelected() {
		var value = this.formContext.getAttribute(this.fieldLogicalName).getValue();
		var relatedEntity = this.rlEntities.find(c => c.LogicalName == value);
		if (relatedEntity) {
			this.formContext.getAttribute('maptaskr_referencingattribute').setValue(relatedEntity.ReferencingAttribute);
			const dataSource = this.datasources.find(d => d.maptaskr_name == relatedEntity.LogicalName);
			this.setAddressFields(dataSource);
			this.setAddressFieldsDisabled(true);
		}
	}

	/**
	 * Get entities 
	 */
	private async getEntities(): Promise<string> {
		const param = "EntityDefinitions?$select=SchemaName,EntitySetName,LogicalName";
		return this.request(param);
	}

	/**
	 * Populate attributes
	 */
	private async populateAttributes(entity: string) {
		this.records = [];
		var a = await this.getAttributes(entity);
		var result = JSON.parse(a);

		for (var i = 0; i < result.value.length; i++) {
			if (result.value[i].LogicalName !== null) {
				this.records.push(result.value[i].LogicalName);
			}
		}
	}

	/**
	 * Retreive attributes from server
	 * @param entity 
	 */
	private async getAttributes(entity: string): Promise<string> {
		const param = "EntityDefinitions(LogicalName='" + entity + "')/Attributes?$select=LogicalName,DisplayName,AttributeType&$filter=AttributeOf%20eq%20null";
		return this.request(param);
	}

	/**
	 * Populate related entities
	 */
	private async populateRelatedEntities(entity: string) {
		this.records = [];

		// Entity metadata
		var metaData = await this.getEntityMetaData(entity);
		var entityMetaData = JSON.parse(metaData);
		var metadataId = entityMetaData.value ? entityMetaData.value[0].MetadataId : undefined;

		if (!metadataId) return;

		// Releted entities
		var result = await this.getRelatedEntities(metadataId);
		var parsedRlEntitiesResult = JSON.parse(result);

		for (var i = 0; i < parsedRlEntitiesResult.value.length; i++) {
			if (parsedRlEntitiesResult.value[i].ReferencedEntity !== null) {
				this.records.push(parsedRlEntitiesResult.value[i].ReferencedEntity);
			}
		}

		var dsResult = await this.getDatasources();
		var dsParsedResult = JSON.parse(dsResult);
		this.datasources = dsParsedResult.value;

		if (this.datasources) this.rlEntities = this.generateRelatedEntities(parsedRlEntitiesResult.value);

	}

	/**
	 * Get related entities from server
	 * @param entity parent
	 */
	private async getRelatedEntities(metadataId: string): Promise<string> {
		const param = "EntityDefinitions(" + metadataId + ")/ManyToOneRelationships?$select=ReferencedEntity,ReferencingAttribute,MetadataId,SchemaName";
		return this.request(param);
	}

	/**
	 * Retreive entity metadata from server
	 * @param entity 
	 */
	private async getEntityMetaData(entity: string): Promise<string> {
		const param = "EntityDefinitions?$select=SchemaName,LogicalName,MetadataId&$filter=LogicalName eq '" + entity + "'";
		return this.request(param);
	}

	/**
	 * Get related entities from server
	 * @param entity parent
	 */
	private async getDatasources(): Promise<string> {
		const param = "maptaskr_datasources";
		return this.request(param);
	}

	/**
	 * Common call back for api
	 * @param param 
	 */
	private request(param: string): Promise<string> {
		var req = new XMLHttpRequest();
		var baseUrl = this.baseUrl;

		return new Promise(function (resolve, reject) {

			req.open("GET", baseUrl + "/api/data/v9.1/" + param, true);
			req.onreadystatechange = function () {

				if (req.readyState !== 4) return;
				if (req.status >= 200 && req.status < 300) {

					// If successful
					try {

						var result = JSON.parse(req.responseText);
						if (parseInt(result.StatusCode) < 0) {
							reject({
								status: result.StatusCode,
								statusText: result.StatusMessage
							});
						}
						resolve(req.responseText);
					}
					catch (error) {
						throw error;
					}

				} else {
					// If failed
					reject({
						status: req.status,
						statusText: req.statusText
					});
				}

			};
			req.setRequestHeader("OData-MaxVersion", "4.0");
			req.setRequestHeader("OData-Version", "4.0");
			req.setRequestHeader("Accept", "application/json");
			req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			req.send();
		});
	}

	/**
	 * Set global form context
	 */
	private setFormContext() {
		this.formContext = (<any>window.parent).pageContextForPCF ? (<any>window.parent).pageContextForPCF : undefined;
	}

	/**
	 * Set parent entity
	 */
	private setParentEntity() {
		this.entity = this.formContext ? this.formContext.getAttribute('maptaskr_name').getValue() : '';
	}

	/**
	 * Set fields on selecting related entity who has address
	 * @param dataSource 
	 */
	private setAddressFields(dataSource: any) {
		this.setFieldValue('maptaskr_addresscomposite', dataSource.maptaskr_addresscomposite);
		this.setFieldValue('maptaskr_street1', dataSource.maptaskr_street1);
		this.setFieldValue('maptaskr_street2', dataSource.maptaskr_street2);
		this.setFieldValue('maptaskr_street3', dataSource.maptaskr_street3);
		this.setFieldValue('maptaskr_city', dataSource.maptaskr_city);
		this.setFieldValue('maptaskr_state', dataSource.maptaskr_state);
		this.setFieldValue('maptaskr_postcode', dataSource.maptaskr_postcode);
		this.setFieldValue('maptaskr_country', dataSource.maptaskr_country);
		this.setFieldValue('maptaskr_email', dataSource.maptaskr_email);
		this.setFieldValue('maptaskr_phonenumber', dataSource.maptaskr_phonenumber);
		this.setFieldValue('maptaskr_longitude', dataSource.maptaskr_longitude);
		this.setFieldValue('maptaskr_latitude', dataSource.maptaskr_latitude);
	}

	/**
	 * Common method to set field value
	 * @param fieldName 
	 * @param fieldValue 
	 */
	private setFieldValue(fieldName: string, fieldValue: string) {
		var control = this.formContext.getAttribute(fieldName);
		if (control) control.setValue(fieldValue);
	}

	/**
	 * Common method to disable fields
	 * @param isDisabled 
	 */
	private setAddressFieldsDisabled(isDisabled: boolean) {
		this.disableField('maptaskr_addresscomposite', isDisabled);
		this.disableField('maptaskr_street1', isDisabled);
		this.disableField('maptaskr_street2', isDisabled);
		this.disableField('maptaskr_street3', isDisabled);
		this.disableField('maptaskr_city', isDisabled);
		this.disableField('maptaskr_state', isDisabled);
		this.disableField('maptaskr_postcode', isDisabled);
		this.disableField('maptaskr_country', isDisabled);
		this.disableField('maptaskr_email', isDisabled);
		this.disableField('maptaskr_phonenumber', isDisabled);
		this.disableField('maptaskr_longitude', isDisabled);
		this.disableField('maptaskr_latitude', isDisabled);
	}

	/**
	 * Common method to make field readonly 
	 * @param fieldName 
	 * @param isDisabled 
	 */
	private disableField(fieldName: string, isDisabled: boolean) {
		var control = this.formContext.getControl(fieldName);
		if (control) control.setDisabled(isDisabled);
	}

	/**
	 * Generate related entities based on parent entity
	 * @param relationalEntities 
	 */
	private generateRelatedEntities(relationalEntities: any[]) {
		var rEntities = [];
		for (var i = 0; i < this.datasources.length; i++) {
			var item = relationalEntities.find(r => r.ReferencedEntity == this.datasources[i].maptaskr_name);
			if (item) {
				var isExist = rEntities.find(r => r.LogicalName == item.ReferencedEntity);
				if (!isExist) {
					rEntities.push({
						MetadataId: item.MetadataId,
						LogicalName: item.ReferencedEntity,
						ReferencingAttribute: item.ReferencingAttribute
					});
				}
			}
		}
		return rEntities;
	}
}