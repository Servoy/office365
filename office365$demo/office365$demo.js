/**
 * @properties={typeid:24,uuid:"CD8E1195-03C5-4E62-B613-FD238BA74DA5"}
 */
function getProductsAsArray() {
	var query = "select productname, unitsinstock from products"

	var ds = databaseManager.getDataSetByQuery("example_data", query, [], -1);
	var matrix = []
	for (var i = 1; i <= ds.getMaxRowIndex(); i++) {
		matrix.push(ds.getRowAsArray(i));

	}
	return matrix;
}