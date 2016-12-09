angular.module('office365Word', ['servoy']).factory("office365Word", ['$services', '$window', '$log', '$q', function($services, $window, $log, $q) {

		// MANIFEST LINE: JS-ClientLibs: //appsforoffice.microsoft.com/lib/1/hosted/office.debug.js;version=2.2;name=office.js
		// JS-ClientLibs: office365/lib/office.js;version=2.2;name=office.js,office365/lib/MicrosoftAjax.js;version=2.2;name=MicrosoftAjax.js,office365/lib/o15apptofilemappingtable.js;version=2.2;name=o15apptofilemappingtable.js

		var scope = $services.getServiceScope('office365Word');
		var Office = parent.Office;
		var Word = parent.Word;

		var CONTENT_TYPE = {
			TEXT: "text",
			HTML: "html",
			OOXML: "ooxml"
		}

		var PARTS = {
			BODY: "body",
			PARAGRAPH: "paragraphs"
		}

		// does it make sense ?
		if (Word) {
			var INSERT_LOCATION = {
				END: Word.InsertLocation.end,
				START: Word.InsertLocation.start,
				REPLACE: Word.InsertLocation.replace
			}
		}

		/**
		 * @param error
		 * @param onError
		 * @param officeResultDeferred
		 *
		 * resolve promis with error
		 *  */
		function resolveError(error, onError, officeResultDeferred) {
			var msg = typeof(error) === "string" ? error : error.message;
			$log.error('office365Word: ', msg);
			if (onError) {
				$window.executeInlineScript(onError.formname, onError.script, [msg]);
			}
			officeResultDeferred.resolve(null);
		}

		/**
		 * @param {String} title
		 * @param {Function} syncCallback
		 * @param {Function} onError
		 * @param officeResultDeferred
		 *  */
		function getContentControlByTitle(title, syncCallback, onError, officeResultDeferred) {
			// Run a batch operation against the Word object model.
			Word.run(function(context) {

				var contentControls = context.document.contentControls.getByTitle(title);
				context.load(contentControls, 'text, title, tag');
				return context.sync().then(function(asyncResult) {
					syncCallback.call(this, contentControls, context);
				});
			}).catch(function(error) {
				resolveError(error, onError, officeResultDeferred);
			});
		}

		return {
			isOfficeEnabled: function() {
				return Office ? true : false;
			},
			getDownloadUrl: function(onError) {
				try {
					return Office.context.document.url
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}
			},
			getDocumentUrl: function(onError) {
				try {
					var result;
					var url = Office.context.document.url;
					if (url.indexOf(".sharepoint.com") > -1) {
						// is a sharepoint
						var baseUrl = url.substr(0, url.indexOf("download.aspx?"))
						var token = url.substr(url.indexOf("%7B"), url.length)
						//var token = wrappedToken.replace("%7B","").replace("%7D","");
						result = baseUrl + "WopiFrame.aspx?sourcedoc=" + token;
					} else { // one drive
						result = url.replace("download.aspx", "edit.aspx");
					}
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}
				return result;
			},
			getSelectedData: function(onError) {

				var officeResultDeferred = $q.defer();
				try {
					Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, getSelectedDataCallback);

					function getSelectedDataCallback(result) {
						if (result.status === Office.AsyncResultStatus.Succeeded) {
							officeResultDeferred.resolve(result.value);
						} else {
							resolveError(result.error, onError, officeResultDeferred);
						}
					}
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			setSelectedData: function(text, type, onError) {

				var officeResultDeferred = $q.defer();

				try {
					Office.context.document.setSelectedDataAsync(text, { coercionType: type }, setSelectedDataCallback);
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				function setSelectedDataCallback(result) {
					var returnValue;
					if (result.status === Office.AsyncResultStatus.Succeeded) {
						returnValue = true;
					} else {
						resolveError(result.error, onError, officeResultDeferred)
					}
					officeResultDeferred.resolve(returnValue);
				}

				return officeResultDeferred.promise;
			},
			getBodyOoxml: function(onError) {
				/**
				 * Returns the document body as a text
				 *  */
				var officeResultDeferred = $q.defer();
				Word.run(function(ctx) {
					// Create a proxy object for the document body.
					var body = ctx.document.body;
					var result = body.getOoxml();

					return ctx.sync().then(function() {
						$log.debug(result.value)
						officeResultDeferred.resolve(result.value);
					});

				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred);
				});

				return officeResultDeferred.promise;
			},
			getBodyHtml: function(onError) {
				/**
				 * Returns the document body as a text
				 *  */
				var officeResultDeferred = $q.defer();
				Word.run(function(ctx) {
					// Create a proxy object for the document body.
					var body = ctx.document.body;
					var result = body.getHtml();

					return ctx.sync().then(function() {
						$log.debug(result.value)
						officeResultDeferred.resolve(result.value);
					});
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred);
				});

				return officeResultDeferred.promise;
			},
			getBodyText: function(loadOptions, onError) {
				/**
				 * Returns the document body as a text
				 *  */
				var officeResultDeferred = $q.defer();

				// default option
				// TODO should allow or options or return just text ?
				if (!loadOptions) {
					loadOptions = 'text';
				}

				Word.run(function(ctx) {
					// Create a proxy object for the document body.
					var body = ctx.document.body;
					ctx.load(body, loadOptions);
					officeResultDeferred.resolve(body);
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred);
				});

				return officeResultDeferred.promise;
			},
			insertTextToBody: function(text, insertLocation, onError) {
				var officeResultDeferred = $q.defer();

				if (!insertLocation) {
					insertLocation = Word.InsertLocation.end;
				}

				Word.run(insertText).catch(function(e) {
					resolveError(e, onError, officeResultDeferred);
				});

				function insertText(ctx) {
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
			insertHtmlToBody: function(htmlText, insertLocation, onError) {
				var officeResultDeferred = $q.defer();

				if (!insertLocation) {
					insertLocation = Word.InsertLocation.end;
				}

				Word.run(insertHtml).catch(function(e) {
					resolveError(e, onError, officeResultDeferred);
				});

				function insertHtml(ctx) {
					// Create a proxy object for the document body.
					ctx.document.body.insertHtml(htmlText, insertLocation);
					return ctx.sync().then(function() { // success
							officeResultDeferred.resolve(true);
						}, function(e) { // error
							resolveError(e, onError, officeResultDeferred);
						});
				}

				return officeResultDeferred.promise;
			},
			insertOoxmlToBody: function(ooxmlText, insertLocation, onError) {
				var officeResultDeferred = $q.defer();

				if (!insertLocation) {
					insertLocation = Word.InsertLocation.end;
				}

				Word.run(insertOoxml).catch(function(e) {
					resolveError(e, onError, officeResultDeferred);
				});

				function insertOoxml(ctx) {
					//var myPar = parent.ctx.document.body.insertOoxml('<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ><w:body><w:p><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD  my-first-field  \* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="00211A63"><w:rPr><w:noProof/></w:rPr><w:t>«my-first-field»</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>', 'End');

					// Create a proxy object for the document body.
					ctx.document.body.insertOoxml(ooxmlText, insertLocation);
					return ctx.sync().then(function() {
							officeResultDeferred.resolve(true);
						}, function(e) {
							resolveError(e, onError, officeResultDeferred);
						});
				}

				return officeResultDeferred.promise;
			},
			insertContentControl: function(tag, ops, onError) {

				var officeResultDeferred = $q.defer();

				// Run a batch operation against the Word object model.
				Word.run(function(context) {

					// Create a proxy range object for the current selection.
					var range = context.document.getSelection();

					// Queue a commmand to create the content control.
					var contentControl = range.insertContentControl();
					contentControl.tag = tag;
					if (ops) {
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
			selectBody: function(selectionMode, onError) {
				var officeResultDeferred = $q.defer();

				Word.run(function(ctx) {
					// Create a proxy object for the document body.
					var body = ctx.document.body;
					body.select(selectionMode);
					return ctx.sync().then(function() {
						officeResultDeferred.resolve(true);
					});
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				})

				return officeResultDeferred.promise;
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
				var officeResultDeferred = $q.defer();

				getContentControlByTitle(title, syncCallback, onError, officeResultDeferred);

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

				function syncCallback(contentControls, context) {
					if (contentControls.items.length === 0) {
						officeResultDeferred.resolve(false);
					} else {
						//						for (var i = 0; i < contentControls.items.length; i++) {
						//							contentControls.items[i].insertText(text, "replace");
						//						}
						var item = contentControls.items[0];
						var id = item.id;
						if (typeof(id) === "number") {
							id = id + "";
						}
						Office.context.document.bindings.addFromNamedItemAsync(item.title, Office.CoercionType.Text, { "id": id }, setBindingCallback);
						
						// 1 set binding to named space
						function setBindingCallback(asyncResult) {
							if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
								Office.select("bindings#" + id, resolveError).setDataAsync(text, setBindingDataCallback);
							} else {
								resolveError(asyncResult.error, onError, officeResultDeferred);
							}
						}
						
						// 2 replace the text of the binding callback
						function setBindingDataCallback(result) {
							if (result.status === Office.AsyncResultStatus.Succeeded) {
								Office.context.document.bindings.releaseByIdAsync(id, releaseByIdCallback);
							} else {
								resolveError(result.error, onError, officeResultDeferred);
							}
						}

						// 3 release the binding callback
						function releaseByIdCallback(asyncResult) {
							$log.debug(asyncResult);
							if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
								// 3 replace the text of the binding
								officeResultDeferred.resolve(true);
							} else {
								resolveError(asyncResult.error, onError, officeResultDeferred);
							}
						}
						
					}
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
			getSearchResults: function() {
				return scope.model.searchResults;
			},
			selectSearchResult: function(index, selectionMode, onError) {

				// search result not available
				if (!scope.model.searchResults || !scope.model.searchResults[index]) {
					return false;
				}

				var searchResult = scope.model.searchResults[index];

				var officeResultDeferred = $q.defer();

				Word.run(function(ctx) {
					searchResult.contentControls.select(selectionMode);
					return ctx.sync().then(function() {
						// store the searchResults in cache
						officeResultDeferred.resolve(true);
					});
				}).catch(function(e) {
					resolveError(e, onError, officeResultDeferred)
				})

				return officeResultDeferred.promise;

			}
		}
	}]).run(function($rootScope, $services) {

	//____logProvider.debugEnabled(true);

	//	requirejs.config({
	//		//By default load any module IDs from js/lib
	//		baseUrl: '/'
	//	});
	//
	//	requirejs(["office365/lib/office.debug.js", "office365/lib/MicrosoftAjax.debug.js"], function(util) {
	//			$log.debug("officejs loaded");
	//		});

});