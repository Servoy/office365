/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"F7C6D30E-A783-42CB-907C-7E9B2B527BF2"}
 */
function sendMessage(event) {
	plugins.office365Word.getSelectedData(getSelectedDataCallback)
}

/**
 * @param result
 *
 * @properties={typeid:24,uuid:"F83A0747-5524-40F0-A088-B26ABE85D031"}
 */
function getSelectedDataCallback(result) {
	application.output(result);
}

/**

 * @protected
 *
 * @properties={typeid:24,uuid:"3A2F3521-09DA-411E-8230-27F1FE096CC3"}
 */
function onSwipeLeft() {
	foundset.setSelectedIndex(foundset.getSelectedIndex() + 1)
}

/**
 * @protected 
 * 
 * @properties={typeid:24,uuid:"B8AA94F4-049E-49C0-9C2B-8A6586D1F783"}
 */
function onSwipeRight() {
	foundset.setSelectedIndex(foundset.getSelectedIndex() - 1)
}
