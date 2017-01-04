/**
 * @type {String}
 *
 * @properties={typeid:35,uuid:"11BC7DD5-8871-47D3-9660-E24674F0FDD2"}
 */
var searchProduct = null;

/**

 * @protected
 *
 * @properties={typeid:24,uuid:"0101AE89-2348-4C6E-B17D-2D5123FBEB81"}
 */
function onSwipeLeft() {
	foundset.setSelectedIndex(foundset.getSelectedIndex() + 1)
}

/**

 * @protected
 *
 * @properties={typeid:24,uuid:"E952BCB8-6C74-4AF6-BE07-786C6F1DDDCD"}
 */
function onSwipeRight() {
	foundset.setSelectedIndex(foundset.getSelectedIndex() - 1)
}

/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"AE877419-9231-447B-AA41-AC4D757B6B76"}
 * @AllowToRunInFind
 */
function onSearch(event) {
	if (searchProduct) {
		foundset.selectRecord(searchProduct)
		// searchProduct  = null;
	}
	return true;
}
