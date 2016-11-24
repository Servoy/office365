angular.module('office365Word', ['servoy']).factory("office365Word", ['$services', '$window', '$log', '$q', function($services, $window, $log, $q) {
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
		 * resolve promis with error
		 *  */
		function resolveError(error, onError, officeResultDeferred) {
			$log.error('office365Word: ', error.message);
			if (onError) {
				$window.executeInlineScript(onError.formname, onError.script, [error.message]);
			}
			officeResultDeferred.resolve(null);
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
			getOoxml: function(onError) {
				/**
				 * Returns the document body as a text
				 *  */
				var officeResultDeferred = $q.defer();
				try {
					Word.run(function(ctx) {
						// Create a proxy object for the document body.
						var body = ctx.document.body;
						var result = body.getOoxml();

						return ctx.sync().then(function() {
							console.log(result.value)
							officeResultDeferred.resolve(result.value);
						});

					});
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			getHtml: function(onError) {
				/**
				 * Returns the document body as a text
				 *  */
				var officeResultDeferred = $q.defer();
				try {
					Word.run(function(ctx) {
						// Create a proxy object for the document body.
						var body = ctx.document.body;
						var result = body.getHtml();

						return ctx.sync().then(function() {
							console.log(result.value)
							officeResultDeferred.resolve(result.value);
						});
					});
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			getText: function(loadOptions, onError) {
				/**
				 * Returns the document body as a text
				 *  */
				var officeResultDeferred = $q.defer();

				// default option
				// TODO should allow or options or return just text ?
				if (!loadOptions) {
					loadOptions = 'text';
				}

				try {
					Word.run(function(ctx) {
						// Create a proxy object for the document body.
						var body = ctx.document.body;
						ctx.load(body, loadOptions);
						officeResultDeferred.resolve(body);
					});
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			insertText: function(text, insertLocation, onError) {
				var officeResultDeferred = $q.defer();

				if (!insertLocation) {
					insertLocation = Word.InsertLocation.end;
				}

				try {
					Word.run(function(ctx) {
						
						// Create a proxy object for the document body.
						var body = ctx.document.body;
						ctx.load(body, 'text');
						ctx.document.body.insertText(text, insertLocation);
						return ctx.sync().then(function() {
							officeResultDeferred.resolve(true);
						});
					});
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			insertOoxml: function(ooxmlText, insertLocation, onError) {
				var officeResultDeferred = $q.defer();

				if (!insertLocation) {
					insertLocation = Word.InsertLocation.end;
				}

				try {
					Word.run(function(ctx) {
						//var myPar = parent.ctx.document.body.insertOoxml('<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ><w:body><w:p><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD  my-first-field  \* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="00211A63"><w:rPr><w:noProof/></w:rPr><w:t>«my-first-field»</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>', 'End');

						// Create a proxy object for the document body.
						ctx.document.body.insertOoxml(ooxmlText, insertLocation);
						return ctx.sync().then(function() {
							officeResultDeferred.resolve(true);
						});
					});
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			select: function(selectionMode, onError) {
				var officeResultDeferred = $q.defer();

				try {
					Word.run(function(ctx) {
						// Create a proxy object for the document body.
						var body = ctx.document.body;
						body.select(selectionMode);
						return ctx.sync().then(function() {
							officeResultDeferred.resolve(true);
						});
					});
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			search: function(searchText, searchOptions, onError) {
				var officeResultDeferred = $q.defer();

				try {
					Word.run(function(ctx) {
						// Create a proxy object for the document body.
						var body = ctx.document.body;
						var searchResults = body.search(searchText, searchOptions);
						ctx.load(searchResults, 'text, id, title, font/size');
						return ctx.sync().then(function() {
							console.log(searchResults)
							officeResultDeferred.resolve(searchResults.items);
						});
					});
				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			getBinding: function(id, onError) {
				var officeResultDeferred = $q.defer();

				try {
					Office.context.document.bindings.getByIdAsync(id, callback);

					function callback(asyncResult) {
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

				} catch (e) {
					resolveError(e, onError, officeResultDeferred)
				}

				return officeResultDeferred.promise;
			},
			getAllBindings: function(onError) {
				var officeResultDeferred = $q.defer();

				try {
					//Go to binding by id.
					Office.context.document.bindings.getAllAsync(callback);

					function callback(asyncResult) {
						if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
							var result = asyncResult.value;
							console.log(result);
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

						// get each data async
						function getDataAsync(data) {
							if (data.status === Office.AsyncResultStatus.Succeeded) {
								console.log(data)
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
			gotoBinding: function(id, selectBinding, onError) {
				var officeResultDeferred = $q.defer();

				try {
					var selectionMode = selectBinding ? Office.SelectionMode.Selected : Office.SelectionMode.None
					//Go to binding by id.
					Office.context.document.goToByIdAsync(id, Office.GoToType.Binding, { selectionMode: selectionMode }, callback);

					// FIXME callback is not executed, i can't confirm if selection is done.
					function callback(asyncResult) {
						console.log(asyncResult)
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
			setBindingData: function(id, text, onError) {

				var officeResultDeferred = $q.defer();

				try {
					// 1 Select the binding by id.
					Office.context.document.goToByIdAsync(id, Office.GoToType.Binding, { selectionMode: Office.SelectionMode.Selected }, callback);

					// FIXME callback is not executed, i can't confirm if selection is done.
					function callback(asyncResult) {
						console.log(asyncResult)
					}

					// 2 release binding
					Office.context.document.bindings.releaseByIdAsync(id, releaseByIdCallback);

				} catch (e) {
					$log.error('office365Word: error setting selected data: ' + e.message);
					$window.executeInlineScript(onError.formname, onError.script, [e.message]);
				}

				// 2 release the binding callback
				function releaseByIdCallback(asyncResult) {
					console.log(asyncResult);
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
			}
		}
	}]).run(function($rootScope, $services) { })