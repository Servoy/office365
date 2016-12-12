/**
 * @param {JSFoundSet} fsSupplier
 *
 * @properties={typeid:24,uuid:"7B3A0FA9-EE56-4B34-B0AE-3FB337FFF88E"}
 */
function showSupplier(fsSupplier) {
	application.output(foundset.loadRecords(fsSupplier));
}
