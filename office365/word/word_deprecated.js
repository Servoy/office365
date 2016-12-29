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
	},
	"insertContentControl" : {
		"parameters": [{
		"name": "contentControlOptions",
		"type": "ContentControl"
	},{
		"name": "onError",
		"type": "function",
		"optional" : true
	}],
	"returns" : "string"
	},
	"insertDisplayTagToBody" : {
	"parameters": [{
			"name": "displayTag",
			"type": "string"
		},{
			"name": "insertLocation",
			"type": "string",
			"optional": true
		},{
			"name": "onError",
			"type": "function",
			"optional" : true
		}],
	"returns" : "boolean"
	},
	"getAllContentControls" : {
	    "parameters": [{
			"name": "onError",
			"type": "function",
			"optional" : true
		}],
	"returns" : "object[]"
	},
	"goToContentControl" : {
	    "parameters": [{
			"name": "title",
			"type": "string"
		},{
			"name": "selectionMode",
			"type": "string",
			"optional": true
		},{
			"name": "onError",
			"type": "function",
			"optional" : true
		}],
	"returns" : "boolean"
	},
	"getAllDisplayTags" : {
	    "parameters": [{
			"name": "onError",
			"type": "function",
			"optional" : true
		}],
	"returns" : "object[]"
	},
	"setContentControlText": {
	"parameters": [{
			"name": "title",
			"type": "string"
		},{
			"name": "text",
			"type": "string"
		},{
			"name": "onError",
			"type": "function",
			"optional" : true
		}],
	"returns" : "int"
	},
	"replaceDisplayTags" : {
	"parameters": [{
			"name": "displayTag",
			"type": "string"
		},{
			"name": "text",
			"type": "string"
		},{
			"name": "onError",
			"type": "function",
			"optional" : true
		}],
	"returns" : "int"
	},
	"deleteContentControl": {
	"parameters": [{
			"name": "title",
			"type": "string"
		},{
			"name": "keepContent",
			"type": "boolean"
		},{
			"name": "onError",
			"type": "function",
			"optional" : true
		}],
	"returns" : "boolean"
	},
	"searchBody": {
    	"parameters": [{
				"name": "searchText",
				"type": "string"
			},{
				"name": "searchOptions",
				"type": "object",
				"optional": true
			},{
				"name": "onError",
				"type": "function",
				"optional" : true
			}],
		"returns" : "object"
	},
	"getSearchResults": {
		"returns" : "ContentControl[]"
	},
	"selectSearchResult" : {
		"parameters": [{
				"name": "index",
				"type": "int"
			},{
				"name": "selectionMode",
				"type": "string",
				"optional": true
			},{
				"name": "onError",
				"type": "function",
				"optional" : true
			}],
		"returns" : "boolean"
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
	},
	searchBody: function(searchText, searchOptions, onError) {
		var officeResultDeferred = $q.defer();

		Word.run(function(ctx) {
			// Create a proxy object for the document body.
			var body = ctx.document.body;
			var searchResults = body.search(searchText, searchOptions);
			// TODO what else to include in the search ?
			if (scope.model.searchResults) {
				ctx.trackedObjects.remove(scope.model.searchResults);
			}
			ctx.load(searchResults, 'text, id, title, font/size');
			return ctx.sync().then(function() {
				// store the searchResults in cache
				$log.debug(searchResults);
				ctx.trackedObjects.add(searchResults);
				scope.model.searchResults = searchResults.items;
				officeResultDeferred.resolve(searchResults.items);
			});
		}).catch(function(e) {
			resolveError(e, onError, officeResultDeferred)
		})

		return officeResultDeferred.promise;
	},
	insertDisplayTagToBody: function(displayTag, insertLocation, onError) {
		var officeResultDeferred = $q.defer();

		if (!insertLocation) {
			insertLocation = Word.InsertLocation.end;
		}

		var text = getDisplayTagToken(displayTag)

		Word.run(insertTextCallback).catch(function(e) {
			resolveError(e, onError, officeResultDeferred);
		});

		function insertTextCallback(ctx) {
			
			var selection = ctx.document.getSelection();
			console.log(selection)
			
			// Create a proxy object for the document body.
			var body = ctx.document.body;
			ctx.load(body, 'text');
			ctx.document.body.insertText(text, insertLocation);
			return ctx.sync().then(function() {
					officeResultDeferred.resolve(true);
				}, function(e) {
					resolveError(e, onError, officeResultDeferred);
				});
		}

		return officeResultDeferred.promise;
	},
	insertContentControl: function(ops, onError) {

		var officeResultDeferred = $q.defer();

		// Run a batch operation against the Word object model.
		Word.run(function(context) {

			// Create a proxy range object for the current selection.
			var range = context.document.getSelection();

			// Queue a commmand to create the content control.
			var contentControl = range.insertContentControl();
			if (ops) {
				if (ops.tag) contentControl.tag = ops.tag;
				if (ops.title) contentControl.title = ops.title;
				if (ops.cannotDelete) contentControl.cannotDelete = ops.cannotDelete;
				if (ops.cannotEdit) contentControl.cannotEdit = ops.cannotEdit;
				if (ops.removeWhenEdited) contentControl.removeWhenEdited = ops.removeWhenEdited;
				if (ops.placeholderText) contentControl.placeholderText = ops.placeholderText;
				if (ops.color) contentControl.color = ops.color;
				if (ops.style) contentControl.style = ops.style;
			}

			// Queue a command to load the id property for the content control you created.
			context.load(contentControl, 'id');

			// Synchronize the document state by executing the queued commands,
			// and return a promise to indicate task completion.
			return context.sync().then(function() {
					var id = contentControl.id;
					if (typeof (id) === "number") {
						id = "" + id;
					}
					officeResultDeferred.resolve(id);
				}, function(e) {
					resolveError(e, onError, officeResultDeferred)
				});
		}).catch(function(e) {
			resolveError(e, onError, officeResultDeferred)
		});

		return officeResultDeferred.promise;
	},
	goToContentControl: function(title, selectionMode, onError) {
		var officeResultDeferred = $q.defer();

		getContentControlByTitle(title, syncCallback, onError, officeResultDeferred)

		function syncCallback(contentControls) {
			if (contentControls.items.length === 0) {
				officeResultDeferred.resolve(false);
			} else {
				var nextItem = contentControls.items[0].select(selectionMode);
				officeResultDeferred.resolve(true);
			}
		}

		return officeResultDeferred.promise;

	},
	getAllContentControls: function(onError) {
		var officeResultDeferred = $q.defer();

		// Run a batch operation against the Word object model.
		Word.run(function(context) {

			var contentControls = context.document.contentControls;
			context.load(contentControls, 'text, title, tag');
			return context.sync().then(function() {
				officeResultDeferred.resolve(contentControls.items);
			});
		}).catch(function(error) {
			resolveError(error, onError, officeResultDeferred);
		});

		return officeResultDeferred.promise;
	},
	setContentControlText: function(title, text, onError) {
		/***
		 * Find all content control having title
		 * Replace the text for each item found
		 * Using binding to replace text is prefered to plain insertText since using binding is possible to mantain the design text style.
		 * insertText will reset the text style to normal instead.
		 * Since there may be several item with the same namesSpace (title),
		 * - select the item
		 * - add binding from selection
		 * - set binding data
		 * - release time
		 *
		 * */

		var officeResultDeferred = $q.defer();

		//getContentControlByTitle(title, syncCallback, onError, officeResultDeferred);

		//				// go to all of them
		//				function syncCallback(contentControls, context) {
		//					if (contentControls.items.length === 0) {
		//						officeResultDeferred.resolve(false);
		//					} else {
		//						for (var i = 0; i < contentControls.items.length; i++) {
		//							contentControls.items[i].insertText(text, "replace");
		//						}
		//						officeResultDeferred.resolve(true);
		//
		//						context.sync().then(function() {
		//							officeResultDeferred.resolve(true);
		//						}).catch(function(e) {
		//							resolveError(e, onError, officeResultDeferred)
		//						})
		//					}
		//				}

		// Run a batch operation against the Word object model.
		Word.run(function(context) {

			var contentControls = context.document.contentControls.getByTitle(title);
			context.load(contentControls, 'text, title, tag');
			// sync
			return context.sync().then(function(asyncResult) {
				// wait until process is complete
				return syncCallback(contentControls, context).then(function(result) {
					officeResultDeferred.resolve(result);
				}).catch(function(error) {
					resolveError(error, onError, officeResultDeferred);
				});
			});
		}).catch(function(error) {
			resolveError(error, onError, officeResultDeferred);
		});

		// 0 search all content control by title
		function syncCallback(contentControls, context) {

			// use promise to force Run method to wait until process is complete
			return new Promise(function(resolveSync, rejectSync) {
				if (contentControls.items.length === 0) {
					resolveSync(0);
				} else {
					var i = 0;
					iterate();

					// iterate over each item of the contentControl
					function iterate() {
						var item = contentControls.items[i];
						if (i < contentControls.items.length) { // until has item to process
							$log.debug('setContentControlTex start ' + item.id);

							i++;
							processItem(item, context).then(function() {
								iterate();
							}).catch(function(error) {
								resolveError(error, onError, officeResultDeferred);
							});

						} else { // process is complete, sync to push all outstanding changes to document and resolve
							return context.sync().then(function() {
								resolveSync(i);
							});
						}
					}
				}
			});
		}

		function processItem(item, context) {

			/* Use promise to make sure that value are processed sequentially */
			return new Promise(function(resolve, reject) {

				var bindingId = Math.random().toString(36).substring(7);

				// 1 select item
				item.select('select')
				return item.context.sync().then(function() {
					$log.debug('setContentControlTex selected ' + item.id)
					var itemId = item.id;
					if (typeof (itemId) === "number") {
						itemId = itemId + "";
					}
					// add binding
					addBindingFromSelection(bindingId);
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				})

				// 2 add binding
				function addBindingFromNameSpace(id) {
					Office.context.document.bindings.addFromNamedItemAsync(title, Office.BindingType.Text, { "id": id, asyncContext: id }, replaceBindingsText);
				}

				function addBindingFromSelection(id) {
					Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { "id": id, asyncContext: id }, replaceBindingsText);
				}

				// 3 replace text
				function replaceBindingsText(asyncRes) {
					console.log(asyncRes)

					$log.debug('setContentControlTex setSelectedBinding ' + item.id)

					var id = asyncRes.asyncContext;

					// 3 set binding to the selected text
					if (asyncRes.status === Office.AsyncResultStatus.Succeeded) {
						asyncRes.value.setDataAsync(text, setBindingDataCallback);
						//Office.select("bindings#" + id, resolveError).setDataAsync(text, setBindingDataCallback);
					} else {
						reject(asyncRes.error);
					}

					// 4 replace the text of the binding callback
					function setBindingDataCallback(result) {
						$log.debug('setContentControlTex replaced ' + item.id)

						if (result.status === Office.AsyncResultStatus.Succeeded) {
							// TODO release the binding anyhow
							Office.context.document.bindings.releaseByIdAsync(bindingId, releaseByIdCallback);
						} else {
							Office.context.document.bindings.releaseByIdAsync(bindingId, { error: result.error }, releaseByIdCallback);
							// reject(result.error);
						}
					}

					// 5 release the binding callback
					function releaseByIdCallback(asyncResult) {

						$log.debug('setContentControlTex released ' + item.id);

						if (asyncResult.asyncContext && asyncResult.asyncContext.error) {
							reject(asyncResult.asyncContext.error);
						}

						$log.debug(asyncResult);
						if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
							// 3 replace the text of the binding
							resolve(true);
						} else {
							reject(asyncResult.error);
						}
					}
				}
			});

		}

		return officeResultDeferred.promise;
	},
	deleteContentControl: function(title, keepContent, onError) {
		var officeResultDeferred = $q.defer();

		getContentControlByTitle(title, syncCallback, onError, officeResultDeferred);

		function syncCallback(contentControls) {
			if (contentControls.items.length === 0) {
				officeResultDeferred.resolve(false);
			} else {
				contentControls.items[0].delete(keepContent);
				officeResultDeferred.resolve(true);
			}
		}

		return officeResultDeferred.promise;
	},
	getAllDisplayTags: function(onError) {
		var officeResultDeferred = $q.defer();

		// Run a batch operation against the Word object model.
		Word.run(function(context) {
			// search for display tags in body
			var body = context.document.body;
			var searchResults = body.search("%%*%%", { matchWildCards: true });
			context.load(searchResults, 'text');
			return context.sync().then(function() {
				var displayTags = searchResults.items.map(function (item){
					 item.text = item.text.substring(2, item.text.length-2);
					 return item;
				});
				
				officeResultDeferred.resolve(displayTags);
			});
		}).catch(function(error) {
			resolveError(error, onError, officeResultDeferred);
		});

		return officeResultDeferred.promise;
	},
	replaceDisplayTags: function(displayTag, text, onError) {

		// TODO to be done in paragraph and footer as well.

		var officeResultDeferred = $q.defer();

		Word.run(function(context) {
			// search for display tags in body
			var searchText = getDisplayTagToken(displayTag);
			var body = context.document.body;
			var searchResults = body.search(searchText, { matchCase: true });
			context.load(searchResults, 'text');
			return context.sync().then(function() {

				return syncCallback(searchResults, context).then(function(result) {
					officeResultDeferred.resolve(result);
				}).catch(function(error) {
					resolveError(error, onError, officeResultDeferred);
				});

			});
		}).catch(function(e) {
			resolveError(e, onError, officeResultDeferred)
		});

		// 0 search all content control by title
		function syncCallback(contentControls, context) {

			// use promise to force Run method to wait until process is complete
			return new Promise(function(resolveSync, rejectSync) {
				if (contentControls.items.length === 0) {
					resolveSync(0);
				} else {
					var i = 0;
					iterate();

					// iterate over each item of the contentControl
					function iterate() {
						var item = contentControls.items[i];
						if (i < contentControls.items.length) { // until has item to process
							$log.debug('replaceDisplayTag ' + item.text);

							i++;
							processItem(item, context).then(function() {
								iterate();
							}).catch(function(error) {
								resolveError(error, onError, officeResultDeferred);
							});

						} else { // process is complete, sync to push all outstanding changes to document and resolve
							return context.sync().then(function() {
								resolveSync(i);
							});
						}
					}
				}
			});
		}

		function processItem(item, context) {

			/* Use promise to make sure that value are processed sequentially */
			return new Promise(function(resolve, reject) {

				var bindingId = Math.random().toString(36).substring(7);

				// 1 select item
				item.select('select');

				// sync the selection
				return item.context.sync().then(function() {
					$log.debug('replaceDisplayTag selected ' + item.text)
					// 2 add binding
					Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { "id": bindingId }, replaceBindingsText);
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				})

				// 3 replace text
				function replaceBindingsText(asyncRes) {

					$log.debug('replaceDisplayTag addBinding ' + item.text)

					// 3 set binding to the selected text
					if (asyncRes.status === Office.AsyncResultStatus.Succeeded) {
						// set the binding in the selected context
						asyncRes.value.setDataAsync(text, setBindingDataCallback);
					} else {
						reject(asyncRes.error);
					}

					// 4 replace the text of the binding callback
					function setBindingDataCallback(result) {
						$log.debug('replaceDisplayTag replaced ' + item.text)

						if (result.status === Office.AsyncResultStatus.Succeeded) {
							Office.context.document.bindings.releaseByIdAsync(bindingId, releaseByIdCallback);
						} else {
							// still the binding has to be released
							Office.context.document.bindings.releaseByIdAsync(bindingId, { error: result.error }, releaseByIdCallback);
						}
					}

					// 5 release the binding callback
					function releaseByIdCallback(asyncResult) {

						$log.debug('replaceDisplayTag binding released ' + item.text);

						// error if setDataAsync has failed
						if (asyncResult.asyncContext && asyncResult.asyncContext.error) {
							reject(asyncResult.asyncContext.error);
						}

						$log.debug(asyncResult);
						if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
							// 3 replace the text of the binding
							resolve(true);
						} else {
							reject(asyncResult.error);
						}
					}
				}
			});

		}

		return officeResultDeferred.promise;
	}
}