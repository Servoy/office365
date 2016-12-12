
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
function onShow(firstShow, event) {

	elements.btn_save_doc.enabled = plugins.office365Word.isOfficeEnabled();
}

/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"B506A47D-3244-4B86-9495-49590919D9A4"}
 */
function gotoTemplate(event) {
	application.showForm(forms.material_template_container);
}

/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"6A7212B4-779D-473D-A7D6-D986D05654D9"}
 */
function btnFillTemplate(event) {
	if (foundset.getSize()) {
		var record = foundset.getSelectedRecord();
		//var bindings = plugins.office365Word.getAllBindings();
		var bindings = plugins.office365Word.getAllContentControls(onError);
		for (var i = 0; i < bindings.length; i++) {
			var binding = bindings[i];
			var value = record[binding.title]
			if (value) {
				//plugins.office365Word.insertBindingToContentControl(binding.title,binding.title,onError);
				//plugins.office365Word.setBindingData(binding.title, value, onError)
				application.output(plugins.office365Word.setContentControlText(binding.title,value,onError))
			}
		}
	}
	
	// replace products
	if (plugins.office365Word.goToContentControl("products","select",onError)) {
		var matrix = scopes.office365.getProductsAsArray();
		plugins.office365Word.setSelectedData(matrix, globals.CTYPE.Matrix, onError);
	}
}


/**
 * TODO generated, please specify type and doc for the params
 * @param e
 *
 * @properties={typeid:24,uuid:"899EFEF3-7AB2-4823-BC3B-6F4A245DAE3A"}
 */
function onError(e) {
	application.output(e);
}