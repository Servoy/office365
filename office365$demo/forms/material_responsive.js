
/**
 * @param {JSFoundSet} fsSupplier
 *
 * @properties={typeid:24,uuid:"5AB6DA19-BF9A-46FE-A915-B5AB51BDCC06"}
 */
function showSupplier(fsSupplier) {
	application.output(foundset.loadRecords(fsSupplier));
}