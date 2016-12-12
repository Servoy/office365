/**
 * @properties={typeid:24,uuid:"91371F7E-7964-41A6-BC20-9164106E191E"}
 */
function getProductsAsArray() {
	var query = "select productname, unitsinstock from products"
	
	var ds = databaseManager.getDataSetByQuery("example_data", query,[],-1);
	var matrix = []
	for (var i = 1; i <= ds.getMaxRowIndex(); i++) {
		matrix.push(ds.getRowAsArray(i));
		
	}
	return matrix;
}