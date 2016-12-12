/**
 * Callback method for when form is shown.
 *
 * @param {Boolean} firstShow form is shown first time after load
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"1C15E8B4-0AF0-4E0F-BBBE-5EE3D444BCF3"}
 */
function onShow(firstShow, event) {
	updateUI();
}

/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @properties={typeid:24,uuid:"82F61A65-53D3-48A4-B143-99463CB6A14E"}
 */
function onOpenFile(event) {
	//plugins.officeservicesWord.
	application.showURL(foundset.downlink)
}

/**
 * @properties={typeid:24,uuid:"AA72A9EE-565D-48DD-BD3C-D635BC9EA28D"}
 * @AllowToRunInFind
 */
function addRecord(url, bodyText) {
	
	var version = 1;
	var name = url;
	foundset.find()
	foundset.downlink = url;	
	foundset.search();
	if (foundset.getSize() > 0) {
		version = foundset.getSize() + 1;
		name = foundset.document_name;
	}
	var newname = plugins.dialogs.showInputDialog("Save File","Enter the name of the document", name);
	if (newname) {
		foundset.newRecord(true);
		foundset.document_version = version;
		foundset.document_name = newname;
		foundset.downlink = url;
		
		databaseManager.saveData();
	}
	
	if (bodyText) {
		if (!foundset.documents_to_doc_templates.getSize()) {
			foundset.documents_to_doc_templates.newRecord();
		}
		foundset.documents_to_doc_templates.body_ooxml = bodyText;
	}
	databaseManager.saveData();

}

/**
 *
 * @param {JSEvent} event
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"02851DA8-BEB6-4289-95EB-F8A5DDF8A51A"}
 */
function onWriteToDocument(event) {
	plugins.office365Word.insertOoxmlToBody(foundset.documents_to_doc_templates.body_ooxml,'end');
}

/**
 * @properties={typeid:24,uuid:"17ABFB0B-52D7-4BDF-90D3-016A67EC1E7D"}
 */
function updateUI() {
	elements.button_writeToDoc.enabled = plugins.office365Word.isOfficeEnabled();
}


