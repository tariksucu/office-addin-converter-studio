/*
 * Runs only in webviews too old for the task pane (Internet Explorer's engine, which Office 2016
 * and 2019 use on older Windows; `nomodule` scripts are skipped by every newer engine). Plain
 * ES5: it only says why the pane cannot work there.
 */
(function () {
  var app = document.getElementById('app');
  if (!app) return;
  while (app.firstChild) app.removeChild(app.firstChild);
  var p = document.createElement('p');
  var text = 'This version of Office is too old for the Converter Studio add-in.';
  p.className = 'notice';
  p.appendChild(document.createTextNode(text + ' Use Office on the web or Microsoft 365.'));
  app.appendChild(p);
})();
