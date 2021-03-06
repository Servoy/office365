
/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @properties={typeid:24,uuid:"FCA068AB-FBCA-4CCF-B538-9F2C17F5E815"}
 */
function onOpenFile(event) {
	application.showURL(foundset.downlink)
}

/**
 * @public 
 * @properties={typeid:24,uuid:"EB568B4A-49AA-4CD7-8546-8B04E8914100"}
 * @AllowToRunInFind
 */
function addRecord(url) {
	
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
	
	databaseManager.saveData();

}
