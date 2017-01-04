
/**
 * @properties={typeid:35,uuid:"5F69E0E3-16C8-4250-B424-22D61045A95A",variableType:-4}
 */
var $supplierid = null;

/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"F6DE2C1F-CE16-40F3-BD23-7502C522A580"}
 */
function sendMessage(event) {
	//plugins.svyToastr.info("Your message has been sent with success to " + foundset.contactname)
}

/**
 * @param {JSEvent} event the event that triggered the action
 *
 * @properties={typeid:24,uuid:"3A759C9F-52B0-4317-AB13-B309B23707F2"}
 */
function insertText(event) {
	if (foundset.getSize() > 0) {
		plugins.office365Word.setSelectedData(foundset.getAddressAsText(), scopes.office365.COERCION_TYPE.TEXT, onError);
	}
}

/**
 * @param e
 *
 * @properties={typeid:24,uuid:"B1FF9FBE-E868-4D23-B3D6-46C1DC9AE2DB"}
 */
function onError(e) {
	application.output(e)
}

/**

 * @protected
 *
 * @properties={typeid:24,uuid:"76C83B9D-15FA-444F-86E2-427650795A33"}
 */
function onSwipeLeft() {
	foundset.setSelectedIndex(foundset.getSelectedIndex() + 1)
}

/**
 * @protected 
 * 
 * @properties={typeid:24,uuid:"BA63CFE2-6480-4936-850D-4841D810FC3C"}
 */
function onSwipeRight() {
	foundset.setSelectedIndex(foundset.getSelectedIndex() - 1)
}

/**
 * Handle changed data, return false if the value should not be accepted. In NGClient you can return also a (i18n) string, instead of false, which will be shown as a tooltip.
 *
 * @param oldValue old value
 * @param newValue new value
 * @param {JSEvent} event the event that triggered the action
 *
 * @return {Boolean}
 *
 * @properties={typeid:24,uuid:"16AF10D9-F47B-4CEC-9CF7-C7DA94408C18"}
 * @AllowToRunInFind
 */
function onDataChangeSupplier(oldValue, newValue, event) {
	if ($supplierid) {
		foundset.find();
		foundset.supplierid = $supplierid;
		foundset.search();
	}
	// $supplierid = null;
	return true;
}

/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @properties={typeid:24,uuid:"6BEB1E98-176B-44FD-9F7D-2CE43106B7B1"}
 */
function onGetTextComponent(event) {
	var result = plugins.office365Word.getSelectedData(scopes.office365.COERCION_TYPE.TEXT,onError);
	if (result) {
		foundset.description = result;
		databaseManager.saveData();
	}
}

/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"5FF671CE-C5B1-40A8-A165-B075E99065AE"}
 */
function btnCancelSearchField(event) {
	$supplierid = null;
}
