/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @properties={typeid:24,uuid:"F851CD09-9AD5-4E80-B9EC-AB514E240990"}
 */
function onSaveFile(event) {
	var url = plugins.office365Word.getDocumentUrl();
	var bodyText = plugins.office365Word.getBodyOoxml();
	if (url) {
		application.output(url)
		forms.material_supplier_documents.addRecord(url, bodyText);
	}
}

/**
 * Callback method for when form is shown.
 *
 * @param {Boolean} firstShow form is shown first time after load
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"85251DBD-9980-438B-8FD9-536427E60B2A"}
 */
function onShow(firstShow, event) {
	elements.btnSave.enabled = plugins.office365Word.isOfficeEnabled();
}
