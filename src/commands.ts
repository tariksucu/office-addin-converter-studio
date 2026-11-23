/**
 * The manifests' FunctionFile. Every ribbon button of these add-ins opens the task pane
 * (ShowTaskpane), so there are no functions to run; Office still loads this page and waits for
 * Office.js to be ready on it.
 */
if (typeof Office !== 'undefined') void Office.onReady(() => undefined);
