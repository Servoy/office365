var spec = {
	"insertBindingToBody": {
		"parameters": [{
			"name": "text",
			"type": "string"
		}, {
			"name": "id",
			"type": "string"
		}, {
			"name": "onError",
			"type": "function",
			"optional": true
		}],
		"returns": "string"
	},
	"insertBindingToContentControl": {
		"parameters": [{
			"name": "title",
			"type": "string"
		}, {
			"name": "id",
			"type": "string"
		}, {
			"name": "onError",
			"type": "function",
			"optional": true
		}],
		"returns": "boolean"
	},
	"getAllBindings": {
		"parameters": [{
			"name": "onError",
			"type": "function",
			"optional": true
		}],
		"returns": "Binding[]"
	},
	"getBinding": {
		"parameters": [{
			"name": "id",
			"type": "string"
		}, {
			"name": "onError",
			"type": "function",
			"optional": true
		}],
		"returns": "Binding"
	},
	"addBindingFromSelection": {
		"parameters": [{
			"name": "bindingType",
			"type": "string"
		}, {
			"name": "id",
			"type": "string",
			"optional": true
		}, {
			"name": "onError",
			"type": "function",
			"optional": true
		}],
		"returns": "string"
	},
	"addBindingFromPrompt": {
		"parameters": [{
			"name": "bindingType",
			"type": "string"
		}, {
			"name": "id",
			"type": "string",
			"optional": true
		}, {
			"name": "prompText",
			"type": "string",
			"optional": true
		}, {
			"name": "onError",
			"type": "function",
			"optional": true
		}],
		"returns": "string"
	},
	"goToBinding": {
		"parameters": [{
			"name": "id",
			"type": "string"
		}, {
			"name": "selectBinding",
			"type": "boolean",
			"optional": true
		}, {
			"name": "onError",
			"type": "function",
			"optional": true
		}],
		"returns": "boolean"
	},
	"releaseBinding": {
		"parameters": [{
			"name": "id",
			"type": "string"
		}, {
			"name": "onError",
			"type": "function",
			"optional": true
		}],
		"returns": "boolean"
	},
	"setBindingData": {
		"parameters": [{
			"name": "id",
			"type": "string"
		}, {
			"name": "text",
			"type": "string"
		}, {
			"name": "onError",
			"type": "function",
			"optional": true
		}],
		"returns": "boolean"
	},
	"setTagText": {
		"parameters": [{
			"name": "tag",
			"type": "string"
		}, {
			"name": "text",
			"type": "string"
		}, {
			"name": "onError",
			"type": "function",
			"optional": true
		}],
		"returns": "boolean"
	}

}

var x = {
	insertBindingToBody: function(text, id, onError) {
		var officeResultDeferred = $q.defer();

		try {
			Office.context.document.setSelectedDataAsync(text, { coercionType: Office.BindingType.Text }, setSelectedDataCallback);
		} catch (e) {
			resolveError(e, onError, officeResultDeferred)
		}

		// 1 insert text
		function setSelectedDataCallback(asyncResult) {
			try {
				if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
					Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: id }, addBindingCallback)
				} else {
					resolveError(asyncResult.error, onError, officeResultDeferred);
				}
			} catch (e) {
				resolveError(e, onError, officeResultDeferred);
			}
		}

		// 2 add binding
		function addBindingCallback(asyncResult) {
			if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
				officeResultDeferred.resolve(asyncResult.value.id);
			} else {
				resolveError(asyncResult.error, onError, officeResultDeferred);
			}
		}

		return officeResultDeferred.promise;
	},

	insertBindingToContentControl: function(title, id, onError) {

		var officeResultDeferred = $q.defer();

		try {
			Office.context.document.bindings.addFromNamedItemAsync(title, Office.CoercionType.Text, { "id": id }, setDataCallback);
		} catch (e) {
			resolveError(e, onError, officeResultDeferred);
		}

		function setDataCallback(asyncResult) {
			if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
				officeResultDeferred.resolve(true);
			} else {
				resolveError(asyncResult.error, onError, officeResultDeferred);
			}
		}

		return officeResultDeferred.promise;
	},
	setTagText: function(tag, text, onError) {
		var officeResultDeferred = $q.defer();

		Word.run(function(ctx) {
			var ccs = ctx.document.contentControls.getByTag(tag);
			ctx.load(ccs, { select: 'text' });
			return ctx.sync().then(function() {
				// Replace the text value for each of the content controls that
				// have a tag called "customer". Highlight the content controls.
				for (var i = 0; i < ccs.items.length; i++) {
					ccs.items[i].insertText(text, "replace");
				}
			})// Synchronize the document state by executing the queued commands.
			.then(ctx.sync).then(function() {
				officeResultDeferred.resolve(true);
			}).catch(function(e) {
				resolveError(e, onError, officeResultDeferred)
			})
		});

		return officeResultDeferred.promise;
	},
	getBinding: function(id, onError) {
		var officeResultDeferred = $q.defer();

		try {
			Office.context.document.bindings.getByIdAsync(id, callback);
		} catch (e) {
			resolveError(e, onError, officeResultDeferred)
		}

		function callback(asyncResult) {
			try {
				if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
					var result = asyncResult.value;
					if (result) {
						result.getDataAsync({ asyncContext: { id: result.id, type: result.type } }, getDataAsync)
					} else {
						officeResultDeferred.resolve(null);
					}
				} else {
					resolveError(asyncResult.error, onError, officeResultDeferred);
				}
			} catch (e) {
				resolveError(e, onError, officeResultDeferred);
			}

			// get binding async
			function getDataAsync(data) {
				var binding = {
					id: id,
					text: data.value,
					type: data.asyncContext.type
				}
				officeResultDeferred.resolve(binding);
			}
		}

		return officeResultDeferred.promise;
	},
	getAllBindings: function(onError) {
		var officeResultDeferred = $q.defer();

		try {
			//Go to binding by id.
			Office.context.document.bindings.getAllAsync(callback);
		} catch (e) {
			resolveError(e, onError, officeResultDeferred)
		}

		function callback(asyncResult) {
			try {
				if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
					var result = asyncResult.value;
					$log.debug(result);
					var count = result.length;
					var bindings = []
					if (count) {
						for (var i = 0; i < count; i++) {
							result[i].getDataAsync({ asyncContext: { id: result[i].id, type: result[i].type } }, getDataAsync)
						}
					} else {
						officeResultDeferred.resolve(bindings);
					}
				} else {
					resolveError(asyncResult.error, onError, officeResultDeferred);
				}
			} catch (e) {
				resolveError(e, onError, officeResultDeferred)
			}

			// get each data async
			function getDataAsync(data) {
				if (data.status === Office.AsyncResultStatus.Succeeded) {
					$log.debug(data)
					var item = {
						id: data.asyncContext.id,
						text: data.value,
						type: data.asyncContext.type
					}
					bindings.push(item)
				} else {
					count--;
				}
				// resolve result when all bindings are retrieved
				if (bindings.length === count) {
					officeResultDeferred.resolve(bindings);
				}
			}
		}

		return officeResultDeferred.promise;
	},
	goToBinding: function(id, selectBinding, onError) {
		var officeResultDeferred = $q.defer();

		try {
			var selectionMode = selectBinding ? Office.SelectionMode.Selected : Office.SelectionMode.None
			//Go to binding by id.
			Office.context.document.goToByIdAsync(id, Office.GoToType.Binding, { selectionMode: selectionMode }, callback);

			// FIXME callback is not executed, i can't confirm if selection is done.
			function callback(asyncResult) {
				$log.debug(asyncResult)
				if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
					officeResultDeferred.resolve(true);
				} else {
					resolveError(asyncResult.error, onError, officeResultDeferred);
				}

			}
			// TODO remove it
			officeResultDeferred.resolve(true);

		} catch (e) {
			resolveError(e, onError, officeResultDeferred)
		}

		return officeResultDeferred.promise;
	},
	releaseBinding: function(id, onError) {
		var officeResultDeferred = $q.defer();

		try {

			Office.context.document.bindings.releaseByIdAsync(id, callback);

			function callback(asyncResult) {
				$log.debug(asyncResult)
				if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
					officeResultDeferred.resolve(true);
				} else {
					resolveError(asyncResult.error, onError, officeResultDeferred);
				}
			}

		} catch (e) {
			resolveError(e, onError, officeResultDeferred)
		}

		return officeResultDeferred.promise;
	},
	addBindingFromSelection: function(bindingType, id, onError) {
		var officeResultDeferred = $q.defer();

		if (!bindingType) {
			bindingType = Office.BindingType.Text;
		}

		try {
			Office.context.document.bindings.addFromSelectionAsync(bindingType, { id: id }, callback)

			function callback(asyncResult) {
				if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
					officeResultDeferred.resolve(asyncResult.value.id);
				} else {
					resolveError(asyncResult.error, onError, officeResultDeferred);
				}
			}

		} catch (e) {
			resolveError(e, onError, officeResultDeferred)
		}

		return officeResultDeferred.promise;
	},
	addBindingFromPrompt: function(bindingType, id, promptText, onError) {
		var officeResultDeferred = $q.defer();

		if (!bindingType) {
			bindingType = Office.BindingType.Text;
		}

		try {
			var options = { }
			if (id) options.id = id;
			if (promptText) options.promptText = promptText;
			Office.context.document.bindings.addFromPromptAsync(bindingType, options, callback);

			function callback(asyncResult) {
				if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
					officeResultDeferred.resolve(asyncResult.value);
				} else {
					resolveError(asyncResult.error, onError, officeResultDeferred);
				}
			}
		} catch (e) {
			resolveError(e, onError, officeResultDeferred)
		}

		return officeResultDeferred.promise;
	},
	setBindingData: function(id, text, onError) {

		var officeResultDeferred = $q.defer();

		try {
			Office.select("bindings#" + id, resolveError).setDataAsync(text, setDataCallback);
		} catch (e) {
			resolveError(e, onError, officeResultDeferred);
		}

		function setDataCallback(asyncResult) {
			if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
				officeResultDeferred.resolve(true);
			} else {
				resolveError(asyncResult.error, onError, officeResultDeferred);
			}
		}

		return officeResultDeferred.promise;

		/*
		 try {
		 // 1 Select the binding by id.
		 Office.context.document.goToByIdAsync(id, Office.GoToType.Binding, { selectionMode: Office.SelectionMode.Selected }, callback);

		 // FIXME callback is not executed, i can't confirm if selection is done.
		 function callback(asyncResult) {
		 $log.debug(asyncResult)
		 }

		 // 2 release binding
		 //Office.context.document.bindings.releaseByIdAsync(id, releaseByIdCallback);

		 } catch (e) {
		 $log.error('office365Word: error setting selected data: ' + e.message);
		 $window.executeInlineScript(onError.formname, onError.script, [e.message]);
		 }

		 // 2 release the binding callback
		 function releaseByIdCallback(asyncResult) {
		 $log.debug(asyncResult);
		 if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
		 // 3 replace the text of the binding
		 Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, setSelectedDataCallback);
		 } else {
		 resolveError(asyncResult.error, onError, officeResultDeferred);
		 }
		 }

		 // 3 replace the text of the binding callback
		 function setSelectedDataCallback(result) {
		 if (result.status === Office.AsyncResultStatus.Succeeded) {
		 officeResultDeferred.resolve(true);
		 } else {
		 resolveError(result.error, onError, officeResultDeferred);
		 }
		 }
		 return officeResultDeferred.promise;
		 */
	}

}