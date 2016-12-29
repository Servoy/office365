
/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @properties={typeid:24,uuid:"44FAFDAB-5A53-4E75-B6B9-269DC47E8B5A"}
 */
function onSaveFile(event) {
	var url = plugins.office365Word.getDocumentUrl();
	if (url) {
		application.output(url)
		forms.material_supplier_documents.addRecord(url);
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
 * @properties={typeid:24,uuid:"3F2B3342-52AA-468C-9663-C34226F6A716"}
 */
function onShow(firstShow, event) {}


/**
 * @param e
 *
 * @properties={typeid:24,uuid:"899EFEF3-7AB2-4823-BC3B-6F4A245DAE3A"}
 */
function onError(e) {
	application.output(e);
}