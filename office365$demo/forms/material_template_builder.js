/**
 * @type {String}
 *
 * @properties={typeid:35,uuid:"8D382248-51AC-4560-B248-645FF352EF63"}
 */
var bindingText = null;

/**
 * @type {String}
 *
 * @properties={typeid:35,uuid:"AFB95BAE-9080-4C50-9B8F-90AA59AEEFE2"}
 */
var docBindings = null;

/**
 * @type {String}
 *
 * @properties={typeid:35,uuid:"060F0669-C739-40BF-9CCB-8CC2013DD88B"}
 */
var dataproviderName = "companyname";

/**
 * Callback method for when form is shown.
 *
 * @param {Boolean} firstShow form is shown first time after load
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"CF767FF9-8FA8-4E9A-AD52-6FEECE27442D"}
 */
function onShow(firstShow, event) {
	if (firstShow) {
		// populate valuelist of dataprovider names
		var table = databaseManager.getTable(foundset);
		var columns = table.getColumnNames();
		columns.push("products");
		application.setValueListItems("dataproviders", columns);
		//		refreshBindings();
	}
}

/**
 * @param {JSEvent} event the event that triggered the action
 *
 * @properties={typeid:24,uuid:"8321F39C-0A4B-453A-95EB-B3DBB85274E7"}
 */
function addBinding(event) {
	//if (bindingText) {
	// application.output(plugins.office365Word.insertBindingToBody(bindingText, dataproviderName, null, onError));
	// application.output(plugins.office365Word.insertContentControl(tag,contentControlOptions,onError))
	//} else {
	/** @type {office365-word.ContentControl} */
	var options = { 
		title: dataproviderName, 
		tag: dataproviderName,
		cannotDelete: false,
		cannotEdit: false
	}
	application.output(plugins.office365Word.insertContentControl(options, onError));
	//}
	refreshBindings();
}

/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"81AF0D98-5DCA-4324-B3B0-5479508F3031"}
 */
function onGetBindings(event) {
	refreshBindings();
}

/**
 * @properties={typeid:24,uuid:"08CC6698-081B-4846-87B3-AEF8670459FC"}
 */
function refreshBindings() {
	var bindings = plugins.office365Word.getAllContentControls(onError);
	var displayValue = [];
	var realValue = [];
	var binding;
	for (var i = 0; bindings && i < bindings.length; i++) {
		binding = bindings[i];
		displayValue.push(binding.title + '-' + binding.text + '-' + binding.id);
		realValue.push(binding.title);
	}
	application.setValueListItems("bindings", displayValue, realValue)
}

/**
 * TODO generated, please specify type and doc for the params
 * @param e
 *
 * @properties={typeid:24,uuid:"0430DAC2-F01D-4091-A130-E57950031BCA"}
 */
function onError(e) {
	application.output(e)
}
/**
 * Handle changed data, return false if the value should not be accepted. In NGClient you can return also a (i18n) string, instead of false, which will be shown as a tooltip.
 *
 * @param {String} oldValue old value
 * @param {String} newValue new value
 * @param {JSEvent} event the event that triggered the action
 *
 * @return {Boolean}
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"D9F71275-8C22-4C12-AE50-47CE60EBE99B"}
 */
function onDataChange(oldValue, newValue, event) {
	//plugins.office365Word.goToBinding(newValue, true, onError);
	plugins.office365Word.goToContentControl(newValue, "select", onError)
	return true
}

/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"C47A42D9-A598-487E-B882-6DC7856E1122"}
 */
function gotoHome(event) {
	application.showForm(forms.material_responsive);
}
