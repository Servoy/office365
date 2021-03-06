/**
 * Perform the element default action.
 *
 * @param {JSEvent} event the event that triggered the action
 *
 * @protected
 *
 * @properties={typeid:24,uuid:"CF3845B4-85EE-42BA-AA2F-03F6187C40C9"}
 */
function onInsertChartData(event) {
	getDataSetAsArray()
}


/**
 * @properties={typeid:24,uuid:"0F69D5F1-628F-4A6A-A90C-19406EEC0AD2"}
 */
function getDataSetAsArray() {
	var matrix = scopes.office365$demo.getProductsAsArray();
	plugins.office365Word.setSelectedData(matrix, scopes.office365.COERCION_TYPE.MATRIX, onError);
}
/**
 * @properties={typeid:24,uuid:"3B75AEC8-1576-440E-9A6E-58F33440ADFE"}
 */
function onError(e) {
	application.output(e)
}