/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ 669
(module) {

module.exports = jQuery;

/***/ }

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	(() => {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = (module) => {
/******/ 			var getter = module && module.__esModule ?
/******/ 				() => (module['default']) :
/******/ 				() => (module);
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/************************************************************************/
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
(() => {

// EXTERNAL MODULE: external "jQuery"
var external_jQuery_ = __webpack_require__(669);
var external_jQuery_default = /*#__PURE__*/__webpack_require__.n(external_jQuery_);
;// ./components/i18n.js
/* i18n integration. This is forked from jarn.jsi18n
 *
 * This is a singleton.
 * Configuration is done on the body tag data-i18ncatalogurl attribute
 *     <body data-i18ncatalogurl="/plonejsi18n">
 *
 *  Or, it'll default to "/plonejsi18n"
 */


var I18N = function I18N() {
  var self = this;
  self.baseUrl = external_jQuery_default()('body').attr('data-i18ncatalogurl');
  self.currentLanguage = external_jQuery_default()('html').attr('lang') || 'en';

  // Fix for country specific languages
  if (self.currentLanguage.split('-').length > 1) {
    self.currentLanguage = self.currentLanguage.split('-')[0] + '_' + self.currentLanguage.split('-')[1].toUpperCase();
  }
  self.storage = null;
  self.catalogs = {};
  self.ttl = 24 * 3600 * 1000;

  // Internet Explorer 8 does not know Date.now() which is used in e.g. loadCatalog, so we "define" it
  if (!Date.now) {
    Date.now = function () {
      return new Date().valueOf();
    };
  }
  try {
    if ('localStorage' in window && window.localStorage !== null && 'JSON' in window && window.JSON !== null) {
      self.storage = window.localStorage;
    }
  } catch (e) {}
  self.configure = function (config) {
    for (var key in config) {
      self[key] = config[key];
    }
  };
  self._setCatalog = function (domain, language, catalog) {
    if (domain in self.catalogs) {
      self.catalogs[domain][language] = catalog;
    } else {
      self.catalogs[domain] = {};
      self.catalogs[domain][language] = catalog;
    }
  };
  self._storeCatalog = function (domain, language, catalog) {
    var key = domain + '-' + language;
    if (self.storage !== null && catalog !== null) {
      self.storage.setItem(key, JSON.stringify(catalog));
      self.storage.setItem(key + '-updated', Date.now());
    }
  };
  self.getUrl = function (domain, language) {
    return self.baseUrl + '?domain=' + domain + '&language=' + language;
  };
  self.loadCatalog = function (domain, language) {
    if (language === undefined) {
      language = self.currentLanguage;
    }
    if (self.storage !== null) {
      var key = domain + '-' + language;
      if (key in self.storage) {
        if (Date.now() - parseInt(self.storage.getItem(key + '-updated'), 10) < self.ttl) {
          var catalog = JSON.parse(self.storage.getItem(key));
          self._setCatalog(domain, language, catalog);
          return;
        }
      }
    }
    if (!self.baseUrl) {
      return;
    }
    external_jQuery_default().getJSON(self.getUrl(domain, language), function (catalog) {
      if (catalog === null) {
        return;
      }
      self._setCatalog(domain, language, catalog);
      self._storeCatalog(domain, language, catalog);
    });
  };
  self.MessageFactory = function (domain, language) {
    language = language || self.currentLanguage;
    return function translate(msgid, keywords) {
      var msgstr;
      if (domain in self.catalogs && language in self.catalogs[domain] && msgid in self.catalogs[domain][language]) {
        msgstr = self.catalogs[domain][language][msgid];
      } else {
        msgstr = msgid;
      }
      if (keywords) {
        var regexp, keyword;
        for (keyword in keywords) {
          if (keywords.hasOwnProperty(keyword)) {
            regexp = new RegExp('\\$\\{' + keyword + '\\}', 'g');
            msgstr = msgstr.replace(regexp, keywords[keyword]);
          }
        }
      }
      return msgstr;
    };
  };
};
/* harmony default export */ const components_i18n = (I18N);
;// ./i18n-wrapper.js


// SENAITE message factory
var t = null;
var i18n_wrapper_t = function _t(msgid, keywords) {
  if (t === null) {
    var i18n = new components_i18n();
    console.debug("*** Loading `senaite.core` i18n MessageFactory ***");
    i18n.loadCatalog("senaite.core");
    t = i18n.MessageFactory("senaite.core");
  }
  return t(msgid, keywords);
};

// Plone message factory
var p = null;
var _p = function _p(msgid, keywords) {
  if (p === null) {
    var i18n = new components_i18n();
    console.debug("*** Loading `plone` i18n MessageFactory ***");
    i18n.loadCatalog("plone");
    p = i18n.MessageFactory("plone");
  }
  return p(msgid, keywords);
};
;// ./components/editform.js
var _excluded = ["name"],
  _excluded2 = ["name"],
  _excluded3 = ["name", "error"],
  _excluded4 = ["message", "level"],
  _excluded5 = ["title", "message"],
  _excluded6 = ["name"],
  _excluded7 = ["name"],
  _excluded8 = ["name", "message"],
  _excluded9 = ["name", "message"],
  _excluded0 = ["name", "value"],
  _excluded1 = ["selector", "html"],
  _excluded10 = ["selector", "name", "value"],
  _excluded11 = ["selector", "event", "name"];
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _slicedToArray(r, e) { return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray(r, e) || _nonIterableRest(); }
function _nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t["return"] && (u = t["return"](), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function _arrayWithHoles(r) { if (Array.isArray(r)) return r; }
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _objectWithoutProperties(e, t) { if (null == e) return {}; var o, r, i = _objectWithoutPropertiesLoose(e, t); if (Object.getOwnPropertySymbols) { var n = Object.getOwnPropertySymbols(e); for (r = 0; r < n.length; r++) o = n[r], -1 === t.indexOf(o) && {}.propertyIsEnumerable.call(e, o) && (i[o] = e[o]); } return i; }
function _objectWithoutPropertiesLoose(r, e) { if (null == r) return {}; var t = {}; for (var n in r) if ({}.hasOwnProperty.call(r, n)) { if (-1 !== e.indexOf(n)) continue; t[n] = r[n]; } return t; }
function _createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t["return"] || t["return"](); } finally { if (u) throw o; } } }; }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _classCallCheck(a, n) { if (!(a instanceof n)) throw new TypeError("Cannot call a class as a function"); }
function _defineProperties(e, r) { for (var t = 0; t < r.length; t++) { var o = r[t]; o.enumerable = o.enumerable || !1, o.configurable = !0, "value" in o && (o.writable = !0), Object.defineProperty(e, _toPropertyKey(o.key), o); } }
function _createClass(e, r, t) { return r && _defineProperties(e.prototype, r), t && _defineProperties(e, t), Object.defineProperty(e, "prototype", { writable: !1 }), e; }
function _toPropertyKey(t) { var i = _toPrimitive(t, "string"); return "symbol" == _typeof(i) ? i : i + ""; }
function _toPrimitive(t, r) { if ("object" != _typeof(t) || !t) return t; var e = t[Symbol.toPrimitive]; if (void 0 !== e) { var i = e.call(t, r || "default"); if ("object" != _typeof(i)) return i; throw new TypeError("@@toPrimitive must return a primitive value."); } return ("string" === r ? String : Number)(t); }
/* SENAITE Edit Form Handler
 *
 * This code handles field changes in edit forms and updates others according to
 * the changes with the help of adapters.
 *
 */

// needed for Bootstrap toasts


// debounce interval (ms) to coalesce rapid DOM mutations before processing
var MUTATION_DEBOUNCE = 100;
var EditForm = /*#__PURE__*/function () {
  function EditForm(config) {
    _classCallCheck(this, EditForm);
    this.config = Object.assign({
      "form_selectors": [],
      "field_selectors": []
    }, config);
    this.hooked_fields = [];

    // queued DOM mutations, processed debounced in `flush_mutations`
    this.mutation_queue = [];
    this.mutation_timer = null;

    // bind event handlers
    this.on_mutated = this.on_mutated.bind(this);
    this.flush_mutations = this.flush_mutations.bind(this);
    this.on_modified = this.on_modified.bind(this);
    this.on_submit = this.on_submit.bind(this);
    this.on_blur = this.on_blur.bind(this);
    this.on_click = this.on_click.bind(this);
    this.on_change = this.on_change.bind(this);
    this.on_reference_select = this.on_reference_select.bind(this);
    this.on_reference_deselect = this.on_reference_deselect.bind(this);
    this.init_forms();
  }

  /**
   * Initialize all form elements given by the config
   */
  return _createClass(EditForm, [{
    key: "init_forms",
    value: function init_forms() {
      var selectors = this.config.form_selectors;
      var _iterator = _createForOfIteratorHelper(selectors),
        _step;
      try {
        for (_iterator.s(); !(_step = _iterator.n()).done;) {
          var selector = _step.value;
          var form = document.querySelector(selector);
          if (form && form.tagName === "FORM") {
            this.setup_form(form);
            this.watch_form(form);
          }
        }
      } catch (err) {
        _iterator.e(err);
      } finally {
        _iterator.f();
      }
    }

    /**
     * Trigger `initialized` event on the form element
     */
  }, {
    key: "setup_form",
    value: function setup_form(form) {
      console.debug("EditForm::setup_form(".concat(form, ")"));
      this.ajax_send(form, {}, "initialized");
    }

    /**
     * Bind event handlers on form fields to monitor changes
     */
  }, {
    key: "watch_form",
    value: function watch_form(form) {
      console.debug("EditForm::watch_form(".concat(form, ")"));
      var fields = this.get_form_fields(form);
      var _iterator2 = _createForOfIteratorHelper(fields),
        _step2;
      try {
        for (_iterator2.s(); !(_step2 = _iterator2.n()).done;) {
          var field = _step2.value;
          this.hook_field(field);
        }
        // observe DOM mutations in form
      } catch (err) {
        _iterator2.e(err);
      } finally {
        _iterator2.f();
      }
      this.observe_mutations(form);
      // bind custom form event handlers
      form.addEventListener("modified", this.on_modified);
      form.addEventListener("mutated", this.on_mutated);
      if (form.hasAttribute("ajax-submit")) {
        form.addEventListener("submit", this.on_submit);
      }
    }

    /**
     * Bind event handlers to field
     */
  }, {
    key: "hook_field",
    value: function hook_field(field) {
      // return immediately if the fields is already hooked
      if (this.hooked_fields.indexOf(field) !== -1) {
        // console.debug(`Field '${field.name}' is already hooked`);
        return;
      }
      if (this.is_button(field) || this.is_input_button(field)) {
        // bind click event
        field.addEventListener("click", this.on_click);
      } else if (this.is_reference(field)) {
        // bind custom events from the ReactJS queryselect widget
        field.addEventListener("select", this.on_reference_select);
        field.addEventListener("deselect", this.on_reference_deselect);
      } else if (this.is_text(field) || this.is_textarea(field)) {
        // bind change event
        field.addEventListener("change", this.on_change);
      } else if (this.is_select(field)) {
        // bind events for select field
        field.addEventListener("click", this.on_click);
        field.addEventListener("change", this.on_change);
      } else if (this.is_radio(field) || this.is_checkbox(field)) {
        // bind click event
        field.addEventListener("click", this.on_click);
      } else {
        // bind blur event
        field.addEventListener("blur", this.on_blur);
      }
      // console.debug(`Hooked field '${field.name}'`);
      // remember hooked fields
      this.hooked_fields = this.hooked_fields.concat(field);
    }

    /**
     * Initialize a DOM mutation observer to rebind dynamic added fields,
     * e.g. for records field etc.
     */
  }, {
    key: "observe_mutations",
    value: function observe_mutations(form) {
      var observer = new MutationObserver(function (mutations) {
        var event = new CustomEvent("mutated", {
          detail: {
            form: form,
            mutations: mutations
          }
        });
        form.dispatchEvent(event);
      });
      // observe the form with all contained elements
      observer.observe(form, {
        childList: true,
        subtree: true,
        attributes: false,
        characterData: false
      });
    }

    /**
     * Handle a single DOM mutation
     */
  }, {
    key: "handle_mutation",
    value: function handle_mutation(form, mutation) {
      var target = mutation.target;
      var parent = target.closest(".field");
      var added = mutation.addedNodes;
      var removed = mutation.removedNodes;
      var selectors = this.config.field_selectors;
      // handle picklist widget
      if (this.is_multiple_select(target)) {
        return this.notify(form, target, "modified");
      }
      // hook new fields, e.g. when the records field "More" button was clicked
      if (added && target.ELEMENT_NODE) {
        var _iterator3 = _createForOfIteratorHelper(target.querySelectorAll(selectors)),
          _step3;
        try {
          for (_iterator3.s(); !(_step3 = _iterator3.n()).done;) {
            var field = _step3.value;
            this.hook_field(field);
          }
        } catch (err) {
          _iterator3.e(err);
        } finally {
          _iterator3.f();
        }
      }
      // notify new added elements, e.g. when a category was expanded or the
      // "show more" button was clicked in listings
      if (added.length > 0) {
        return this.notify_added(form, added, "added");
      }
    }

    /**
     * toggles the submit button
     */
  }, {
    key: "toggle_submit",
    value: function toggle_submit(form, toggle) {
      var btn = form.querySelector("input[type='submit']");
      btn.disabled = !toggle;
    }

    /**
     * toggles the display of the field with the `d-none` class
     */
  }, {
    key: "toggle_field_visibility",
    value: function toggle_field_visibility(field) {
      var toggle = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : true;
      var parent = field.closest(".field");
      var css_class = "d-none";
      if (toggle === false) {
        parent.classList.add(css_class);
      } else {
        parent.classList.remove(css_class);
      }
    }

    /**
     * check if fields have errors
     */
  }, {
    key: "has_field_errors",
    value: function has_field_errors(form) {
      var fields_with_errors = form.querySelectorAll(".is-invalid");
      if (fields_with_errors.length > 0) {
        return true;
      }
      return false;
    }

    /**
    * check field type for text control
    */
  }, {
    key: "is_text_control",
    value: function is_text_control(field) {
      var tagName = field.tagName.toLowerCase();
      var fieldType = field.type.toLowerCase();
      var textTypes = ["text", "search", "tel", "url", "email", "password", "date", "month", "week", "time", "datetime-local", "number"];
      var isTextControl = tagName === "input" && textTypes.includes(fieldType);
      return isTextControl || tagName === "textarea";
    }

    /**
     * set field readonly
     */
  }, {
    key: "set_field_readonly",
    value: function set_field_readonly(form, field) {
      var message = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : null;
      // Only text controls can be made read-only
      if (this.is_text_control(field)) {
        field.setAttribute("readonly", "");
      } else {
        // since for other controls (such as checkboxes and buttons)
        // there is no useful distinction between being
        // read-only and being disabled.
        var fieldName = field.name;
        var hiddenField = document.createElement("input");
        hiddenField.setAttribute("type", "hidden");
        hiddenField.setAttribute("name", fieldName);
        hiddenField.setAttribute("value", field.value);
        field.setAttribute("disabled", "");
        field.setAttribute("name", "disabled-" + fieldName);
        // insert hidden control to first positions into form for search by name
        form.prepend(hiddenField);
      }
      var existing_message = field.parentElement.querySelector("div.message");
      if (existing_message) {
        existing_message.innerHTML = _t(message);
      } else {
        var div = document.createElement("div");
        div.className = "message text-secondary small";
        div.innerHTML = _t(message);
        field.parentElement.appendChild(div);
      }
    }

    /**
     * set field editable
     */
  }, {
    key: "set_field_editable",
    value: function set_field_editable(form, field) {
      var message = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : null;
      // Only text controls can be made read-only
      var formField = field;
      if (this.is_text_control(field)) {
        field.removeAttribute("readonly");
      } else {
        // since for other controls (such as checkboxes and buttons)
        // there is no useful distinction between being
        // read-only and being disabled. We cover other controls like select here.

        var fieldName = field.name;
        form.removeChild(field);
        var disabledFieldName = "disabled-" + fieldName;
        var disabledField = this.get_form_field_by_name(form, disabledFieldName);
        if (disabledField) {
          formField = disabledField;
          disabledField.removeAttribute("disabled");
          disabledField.setAttribute("name", fieldName);
        }
      }
      var existing_message = formField.parentElement.querySelector("div.message");
      if (existing_message) {
        existing_message.innerHTML = _t(message);
      } else {
        var div = document.createElement("div");
        div.className = "message text-secondary small";
        div.innerHTML = _t(message);
        formField.parentElement.appendChild(div);
      }
    }

    /**
     * set field error
     */
  }, {
    key: "set_field_error",
    value: function set_field_error(field, message) {
      field.classList.add("is-invalid");
      var existing_message = field.parentElement.querySelector("div.invalid-feedback");
      if (existing_message) {
        existing_message.innerHTML = _t(message);
      } else {
        var div = document.createElement("div");
        div.className = "invalid-feedback";
        div.innerHTML = _t(message);
        field.parentElement.appendChild(div);
      }
    }

    /**
     * remove field error
     */
  }, {
    key: "remove_field_error",
    value: function remove_field_error(field) {
      field.classList.remove("is-invalid");
      var msg = field.parentElement.querySelector(".invalid-feedback");
      if (msg) {
        msg.remove();
      }
    }

    /**
     * add a status message
     * @param {string} message the message to display in the alert
     * @param {string} level   one of "info", "success", "warning", "danger"
     * @param {object} options additional options to control the behavior
     *                 - option {string} title: alert title in bold
     *                 - option {string} flush: remove previous alerts
     */
  }, {
    key: "add_statusmessage",
    value: function add_statusmessage(message) {
      var level = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "info";
      var options = arguments.length > 2 ? arguments[2] : undefined;
      options = options || {};
      var el = document.createElement("div");
      var title = options.title || "".concat(level.charAt(0).toUpperCase() + level.slice(1));
      el.innerHTML = "\n      <div class=\"alert alert-".concat(level, " alert-dismissible fade show\" role=\"alert\">\n        <strong>").concat(title, "</strong>\n        ").concat(_t(message), "\n        <button type=\"button\" class=\"close\" data-dismiss=\"alert\" aria-label=\"Close\">\n          <span aria-hidden=\"true\">&times;</span>\n        </button>\n      </div>\n    ");
      el = el.firstElementChild;
      var parent = document.getElementById("viewlet-above-content");

      // clear put previous alerts
      if (options.flush) {
        var _iterator4 = _createForOfIteratorHelper(parent.querySelectorAll(".alert")),
          _step4;
        try {
          for (_iterator4.s(); !(_step4 = _iterator4.n()).done;) {
            var _el = _step4.value;
            _el.remove();
          }
        } catch (err) {
          _iterator4.e(err);
        } finally {
          _iterator4.f();
        }
      }
      parent.appendChild(el);
      return el;
    }

    /**
     * add a notification message
     */
  }, {
    key: "add_notification",
    value: function add_notification(title, message, options) {
      options = options || {};
      options = Object.assign({
        animation: true,
        autohide: true,
        delay: 5000
      }, options);
      var el = document.createElement("div");
      el.innerHTML = "\n      <div class=\"toast\" style=\"width:300px\" role=\"alert\"\n           data-animation=\"".concat(options.animation, "\"\n           data-autohide=\"").concat(options.autohide, "\"\n           data-delay=\"").concat(options.delay, "\">\n        <div class=\"toast-header\">\n          <strong class=\"mr-auto\">").concat(title.charAt(0).toUpperCase() + title.slice(1), "</strong>\n          <button type=\"button\" class=\"ml-2 mb-1 close\" data-dismiss=\"toast\" aria-label=\"Close\">\n            <span aria-hidden=\"true\">&times;</span>\n          </button>\n        </div>\n        <div class=\"toast-body\">\n          ").concat(_t(message), "\n        </div>\n      </div>\n    ");
      el = el.firstElementChild;
      var parent = document.querySelector(".toast-container");
      if (!parent) {
        parent = document.createElement("div");
        parent.innerHTML = "\n        <div style=\"position: fixed; top: 0px; right: 0px; width=100%;\">\n          <div class=\"toast-container\" style=\"position: absolute; top: 10px; right: 10px;\">\n          </div>\n        </div>\n      ";
        var wrapper = document.querySelector(".container-fluid");
        wrapper.appendChild(parent);
        parent = parent.querySelector(".toast-container");
      }
      parent.appendChild(el);
      return el;
    }

    /**
     * update the form with the response from the server
     */
  }, {
    key: "update_form",
    value: function update_form(form, data) {
      var _this = this;
      console.info("*** UPDATE FORM ***", data);
      if (data === null) {
        data = {};
      }
      var hide = data.hide || [];
      var show = data.show || [];
      var readonly = data.readonly || [];
      var editable = data.editable || [];
      var errors = data.errors || [];
      var messages = data.messages || [];
      var notifications = data.notifications || [];
      var updates = data.updates || [];
      var html = data.html || [];
      var attributes = data.attributes || [];
      var callbacks = data.callbacks || [];
      var states = data.states || [];
      var listings = data.listings || [];

      // ReactJS widget states
      var _iterator5 = _createForOfIteratorHelper(states),
        _step5;
      try {
        for (_iterator5.s(); !(_step5 = _iterator5.n()).done;) {
          var record = _step5.value;
          var name = void 0,
            rest = void 0;
          var _record = record;
          name = _record.name;
          rest = _objectWithoutProperties(_record, _excluded);
          _record;
          if (name in window.senaite.core.widgets) {
            var widget = window.senaite.core.widgets[name];
            widget.clear_results();
            widget.flush();
            widget.setState(rest);
          }
        }

        // ReactJS listing states
      } catch (err) {
        _iterator5.e(err);
      } finally {
        _iterator5.f();
      }
      var _iterator6 = _createForOfIteratorHelper(listings),
        _step6;
      try {
        for (_iterator6.s(); !(_step6 = _iterator6.n()).done;) {
          var _record2 = _step6.value;
          var _name = void 0,
            _rest = void 0;
          var _record3 = _record2;
          _name = _record3.name;
          _rest = _objectWithoutProperties(_record3, _excluded2);
          _record3;
          if (_name in window.listings) {
            var listing = window.listings[_name] || window.senaite.core.listings[_name];
            listing.setState(_rest);
          }
        }

        // render field errors
      } catch (err) {
        _iterator6.e(err);
      } finally {
        _iterator6.f();
      }
      var _iterator7 = _createForOfIteratorHelper(errors),
        _step7;
      try {
        for (_iterator7.s(); !(_step7 = _iterator7.n()).done;) {
          var _record4 = _step7.value;
          var _name2 = void 0,
            error = void 0,
            _rest2 = void 0;
          var _record5 = _record4;
          _name2 = _record5.name;
          error = _record5.error;
          _rest2 = _objectWithoutProperties(_record5, _excluded3);
          _record5;
          var el = this.get_form_field_by_name(form, _name2);
          if (!el) continue;
          if (error) {
            this.set_field_error(el, error);
          } else {
            this.remove_field_error(el);
          }
        }

        // render status messages
      } catch (err) {
        _iterator7.e(err);
      } finally {
        _iterator7.f();
      }
      var _iterator8 = _createForOfIteratorHelper(messages),
        _step8;
      try {
        for (_iterator8.s(); !(_step8 = _iterator8.n()).done;) {
          var _record6 = _step8.value;
          var _name3 = void 0,
            _error = void 0,
            _rest3 = void 0;
          var _record7 = _record6;
          message = _record7.message;
          level = _record7.level;
          _rest3 = _objectWithoutProperties(_record7, _excluded4);
          _record7;
          var level = level || "info";
          var message = message || "";
          this.add_statusmessage(message, level, _rest3);
        }

        // render notification messages
      } catch (err) {
        _iterator8.e(err);
      } finally {
        _iterator8.f();
      }
      var _iterator9 = _createForOfIteratorHelper(notifications),
        _step9;
      try {
        for (_iterator9.s(); !(_step9 = _iterator9.n()).done;) {
          var _record8 = _step9.value;
          var title = void 0,
            _message = void 0,
            _rest4 = void 0;
          var _record9 = _record8;
          title = _record9.title;
          _message = _record9.message;
          _rest4 = _objectWithoutProperties(_record9, _excluded5);
          _record9;
          var _el2 = this.add_notification(title, _message, _rest4);
          external_jQuery_default()(_el2).toast("show");
        }

        // hide fields
      } catch (err) {
        _iterator9.e(err);
      } finally {
        _iterator9.f();
      }
      var _iterator0 = _createForOfIteratorHelper(hide),
        _step0;
      try {
        for (_iterator0.s(); !(_step0 = _iterator0.n()).done;) {
          var _record0 = _step0.value;
          var _name4 = void 0,
            _rest5 = void 0;
          var _record1 = _record0;
          _name4 = _record1.name;
          _rest5 = _objectWithoutProperties(_record1, _excluded6);
          _record1;
          var _el3 = this.get_form_field_by_name(form, _name4);
          if (!_el3) continue;
          this.toggle_field_visibility(_el3, false);
        }

        // show fields
      } catch (err) {
        _iterator0.e(err);
      } finally {
        _iterator0.f();
      }
      var _iterator1 = _createForOfIteratorHelper(show),
        _step1;
      try {
        for (_iterator1.s(); !(_step1 = _iterator1.n()).done;) {
          var _record10 = _step1.value;
          var _name5 = void 0,
            _rest6 = void 0;
          var _record11 = _record10;
          _name5 = _record11.name;
          _rest6 = _objectWithoutProperties(_record11, _excluded7);
          _record11;
          var _el4 = this.get_form_field_by_name(form, _name5);
          if (!_el4) continue;
          this.toggle_field_visibility(_el4, true);
        }

        // readonly fields
      } catch (err) {
        _iterator1.e(err);
      } finally {
        _iterator1.f();
      }
      var _iterator10 = _createForOfIteratorHelper(readonly),
        _step10;
      try {
        for (_iterator10.s(); !(_step10 = _iterator10.n()).done;) {
          var _record12 = _step10.value;
          var _name6 = void 0,
            _message2 = void 0,
            _rest7 = void 0;
          var _record13 = _record12;
          _name6 = _record13.name;
          _message2 = _record13.message;
          _rest7 = _objectWithoutProperties(_record13, _excluded8);
          _record13;
          var _el5 = this.get_form_field_by_name(form, _name6);
          if (!_el5) continue;
          this.set_field_readonly(form, _el5, _message2);
        }

        // editable fields
      } catch (err) {
        _iterator10.e(err);
      } finally {
        _iterator10.f();
      }
      var _iterator11 = _createForOfIteratorHelper(editable),
        _step11;
      try {
        for (_iterator11.s(); !(_step11 = _iterator11.n()).done;) {
          var _record14 = _step11.value;
          var _name7 = void 0,
            _message3 = void 0,
            _rest8 = void 0;
          var _record15 = _record14;
          _name7 = _record15.name;
          _message3 = _record15.message;
          _rest8 = _objectWithoutProperties(_record15, _excluded9);
          _record15;
          var _el6 = this.get_form_field_by_name(form, _name7);
          if (!_el6) continue;
          this.set_field_editable(form, _el6, _message3);
        }

        // updated fields
      } catch (err) {
        _iterator11.e(err);
      } finally {
        _iterator11.f();
      }
      var _iterator12 = _createForOfIteratorHelper(updates),
        _step12;
      try {
        for (_iterator12.s(); !(_step12 = _iterator12.n()).done;) {
          var _record16 = _step12.value;
          var _name8 = void 0,
            value = void 0,
            _rest9 = void 0;
          var _record17 = _record16;
          _name8 = _record17.name;
          value = _record17.value;
          _rest9 = _objectWithoutProperties(_record17, _excluded0);
          _record17;
          var _el7 = this.get_form_field_by_name(form, _name8);
          if (!_el7) continue;
          this.set_field_value(_el7, value);
        }

        // html
      } catch (err) {
        _iterator12.e(err);
      } finally {
        _iterator12.f();
      }
      var _iterator13 = _createForOfIteratorHelper(html),
        _step13;
      try {
        for (_iterator13.s(); !(_step13 = _iterator13.n()).done;) {
          var _record18 = _step13.value;
          var selector = void 0,
            _html = void 0,
            _rest0 = void 0;
          var _record19 = _record18;
          selector = _record19.selector;
          _html = _record19.html;
          _rest0 = _objectWithoutProperties(_record19, _excluded1);
          _record19;
          var _el8 = form.querySelector(selector);
          if (!_el8) continue;
          if (_rest0.append) {
            _el8.innerHTML = _el8.innerHTML + _html;
          } else {
            _el8.innerHTML = _html;
          }
        }

        // set attribute to an element
      } catch (err) {
        _iterator13.e(err);
      } finally {
        _iterator13.f();
      }
      var _iterator14 = _createForOfIteratorHelper(attributes),
        _step14;
      try {
        for (_iterator14.s(); !(_step14 = _iterator14.n()).done;) {
          var _record20 = _step14.value;
          var _selector = void 0,
            _name9 = void 0,
            _value = void 0,
            _rest1 = void 0;
          var _record21 = _record20;
          _selector = _record21.selector;
          _name9 = _record21.name;
          _value = _record21.value;
          _rest1 = _objectWithoutProperties(_record21, _excluded10);
          _record21;
          var _el9 = form.querySelector(_selector);
          if (!_el9) continue;
          if (_value === null) {
            _el9.removeAttribute(_name9);
          } else {
            _el9.addAttribute(_name9, _value);
          }
        }

        // register callbacks
      } catch (err) {
        _iterator14.e(err);
      } finally {
        _iterator14.f();
      }
      var _iterator15 = _createForOfIteratorHelper(callbacks),
        _step15;
      try {
        var _loop = function _loop() {
          var record = _step15.value;
          var selector, event, name, rest;
          // register local callback to apply additional data
          var _record22 = record;
          selector = _record22.selector;
          event = _record22.event;
          name = _record22.name;
          rest = _objectWithoutProperties(_record22, _excluded11);
          _record22;
          var on_callback = function on_callback(event) {
            console.debug("EditForm::on_callback");
            var data = {
              name: name,
              target: event.currentTarget.name || null,
              value: event.currentTarget.value || null
            };
            _this.ajax_send(form, data, "callback");
          };
          if (selector === "document") {
            document.addEventListener(event, on_callback);
          } else {
            document.querySelectorAll(selector).forEach(function (el) {
              el.addEventListener(event, on_callback);
            });
          }
        };
        for (_iterator15.s(); !(_step15 = _iterator15.n()).done;) {
          _loop();
        }

        // disallow submit when field errors are present
      } catch (err) {
        _iterator15.e(err);
      } finally {
        _iterator15.f();
      }
      if (this.has_field_errors(form)) {
        this.toggle_submit(form, false);
      } else {
        this.toggle_submit(form, true);
      }
    }

    /**
     * return a form field by name
     */
  }, {
    key: "get_form_field_by_name",
    value: function get_form_field_by_name(form, name) {
      // get the first element that matches the name
      var exact = form.querySelector("[name='".concat(name, "']"));
      var fuzzy = form.querySelector("[name^='".concat(name, "']"));
      var field = exact || fuzzy || null;
      if (field === null) {
        return null;
      }
      return field;
    }

    /**
     * return a dictionary of all the form values
     */
  }, {
    key: "get_form_data",
    value: function get_form_data(form) {
      var data = {};
      var form_data = new FormData(form);
      form_data.forEach(function (value, key) {
        data[key] = value;
      });
      // handle DX add views
      var view_name = this.get_view_name();
      if (view_name.indexOf("++add++") > -1) {
        // inject `form_adapter_name` for named multi adapter lookup
        // see senaite.core.browser.form.ajax.FormView for lookup logic
        data.form_adapter_name = view_name;
      }
      return data;
    }

    /**
     * Return form fields for the given selectors of the config
     */
  }, {
    key: "get_form_fields",
    value: function get_form_fields(form) {
      console.debug("EditForm::get_form_fields(".concat(form, ")"));
      var fields = [];
      var selectors = this.config.field_selectors;
      var _iterator16 = _createForOfIteratorHelper(selectors),
        _step16;
      try {
        for (_iterator16.s(); !(_step16 = _iterator16.n()).done;) {
          var _fields;
          var selector = _step16.value;
          var nodes = form.querySelectorAll(selector);
          fields = (_fields = fields).concat.apply(_fields, _toConsumableArray(nodes.values()));
        }
      } catch (err) {
        _iterator16.e(err);
      } finally {
        _iterator16.f();
      }
      return fields;
    }

    /**
     * returns the name of the field w/o ZPublisher converter
     */
  }, {
    key: "get_field_name",
    value: function get_field_name(field) {
      var name = field.name;
      return name.split(":")[0];
    }

    /**
     * return the value of the form field
     */
  }, {
    key: "get_field_value",
    value: function get_field_value(field) {
      if (this.is_checkbox(field)) {
        // returns true/false for checkboxes
        return field.checked;
      } else if (this.is_select(field)) {
        // returns a list of selected option
        var selected = field.selectedOptions;
        return Array.prototype.map.call(selected, function (option) {
          return option.value;
        });
      } else if (this.is_reference(field)) {
        return field.value.split("\n");
      }
      // return the plain field value
      return field.value;
    }

    /**
     * set the value of the form field
     */
  }, {
    key: "set_field_value",
    value: function set_field_value(field, value) {
      // for reference/select fields
      var selected = value.selected || [];
      var options = value.options || [];

      // set reference value
      if (this.is_reference(field)) {
        // Fallback: Use raw value if selected is not set
        if (value && selected.length == 0) {
          selected = value.split("\n");
        }
        // XXX: does not work for ReactJS components!
        // field.value = selected.join("\n");
        this.native_set_value(field, selected.join("\n"));
      }
      // set select field
      else if (this.is_select(field)) {
        if (selected.length == 0) {
          var old_selected = field.options[field.selected];
          if (old_selected) {
            selected = [old_selected.value];
          }
        }
        // remove all options
        field.options.length = 0;
        // sort options
        options.sort(function (a, b) {
          var _a = a.title.toLowerCase();
          var _b = b.title.toLowerCase();
          if (a.value === null) _a = "";
          if (b.value === null) _b = "";
          if (_a < _b) return -1;
          if (_a > _b) return 1;
        });
        // build new options
        var _iterator17 = _createForOfIteratorHelper(options),
          _step17;
        try {
          for (_iterator17.s(); !(_step17 = _iterator17.n()).done;) {
            var option = _step17.value;
            var el = document.createElement("option");
            el.value = option.value;
            el.innerHTML = option.title;
            // select item if the value is in the selected array
            if (selected.indexOf(option.value) !== -1) {
              el.selected = true;
            }
            field.appendChild(el);
          }
          // select first item
        } catch (err) {
          _iterator17.e(err);
        } finally {
          _iterator17.f();
        }
        if (selected.length == 0) {
          field.selectedIndex = 0;
        }
      }
      // set checkbox value
      else if (this.is_checkbox(field)) {
        field.checked = value;
      }
      // set other field values
      else {
        field.value = value;
      }
    }

    /**
     * set input value with native setter to support ReactJS components
     *
     * https://stackoverflow.com/questions/23892547/what-is-the-best-way-to-trigger-onchange-event-in-react-js
     * TL;DR: React library overrides input value setter
     */
  }, {
    key: "native_set_value",
    value: function native_set_value(input, value) {
      var setter = null;
      if (input.tagName === "TEXTAREA") {
        var _Object$getOwnPropert;
        setter = (_Object$getOwnPropert = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value")) === null || _Object$getOwnPropert === void 0 ? void 0 : _Object$getOwnPropert.set;
      } else if (input.tagName === "SELECT") {
        var _Object$getOwnPropert2;
        setter = (_Object$getOwnPropert2 = Object.getOwnPropertyDescriptor(window.HTMLSelectElement.prototype, "value")) === null || _Object$getOwnPropert2 === void 0 ? void 0 : _Object$getOwnPropert2.set;
      } else if (input.tagName === "INPUT") {
        var _Object$getOwnPropert3;
        setter = (_Object$getOwnPropert3 = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value")) === null || _Object$getOwnPropert3 === void 0 ? void 0 : _Object$getOwnPropert3.set;
      } else {
        input.value = value;
      }
      if (setter) {
        setter.call(input, value);
      }
      var event = new Event("input", {
        bubbles: true
      });
      input.dispatchEvent(event);
    }

    /**
     * trigger `modified` event on the form
     */
  }, {
    key: "modified",
    value: function modified(el) {
      var event = new CustomEvent("modified", {
        detail: {
          field: el,
          form: el.form
        }
      });
      // dispatch the event on the element
      el.form.dispatchEvent(event);
    }

    /**
     * trigger ajax loading events
     */
  }, {
    key: "loading",
    value: function loading() {
      var toggle = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : true;
      var event_type = toggle ? "ajaxStart" : "ajaxStop";
      var event = new CustomEvent(event_type);
      document.dispatchEvent(event);
    }

    /**
     * notify a field change to the server ajax endpoint
     */
  }, {
    key: "notify",
    value: function notify(form, field, endpoint) {
      var data = {
        name: this.get_field_name(field),
        value: this.get_field_value(field)
      };
      this.ajax_send(form, data, endpoint);
    }

    /**
     * notify that DOM Elements were added
     */
  }, {
    key: "notify_added",
    value: function notify_added(form, added, endpoint) {
      var data = {
        added: []
      };
      added.forEach(function (el) {
        var record = {};
        if (el.attributes && el.attributes.length > 0) {
          var _iterator18 = _createForOfIteratorHelper(el.attributes),
            _step18;
          try {
            for (_iterator18.s(); !(_step18 = _iterator18.n()).done;) {
              var attribute = _step18.value;
              var name = attribute.name;
              var value = attribute.value;
              record[name] = value;
            }
          } catch (err) {
            _iterator18.e(err);
          } finally {
            _iterator18.f();
          }
        }
        if (Object.keys(record).length > 0) {
          data.added.push(record);
        }
      });
      this.ajax_send(form, data, endpoint);
    }

    /**
     * send application/json to the server
     */
  }, {
    key: "ajax_send",
    value: function ajax_send(form, data, endpoint) {
      var view_url = document.body.dataset.viewUrl;
      var ajax_url = "".concat(view_url, "/ajax_form/").concat(endpoint);
      var payload = Object.assign({
        form: this.get_form_data(form)
      }, data);
      console.debug("EditForm::ajax_send --> ", payload);
      var init = {
        method: "POST",
        credentials: "include",
        body: JSON.stringify(payload),
        headers: {
          "Content-Type": "application/json",
          "X-CSRF-TOKEN": document.querySelector("#protect-script").dataset.token
        }
      };
      return this.ajax_request(form, ajax_url, init);
    }

    /**
     * send multipart/form-data to the server
     *
     * NOTE: This is used by the import form and hooked by the `ajax-submit="1"` attribute
     *       Therefore, we send here the data as multipart/form-data
     */
  }, {
    key: "ajax_submit",
    value: function ajax_submit(form, data, endpoint) {
      var view_url = document.body.dataset.viewUrl;
      var ajax_url = "".concat(view_url, "/ajax_form/").concat(endpoint);
      var payload = new FormData(form);

      // update form data
      for (var _i = 0, _Object$entries = Object.entries(data); _i < _Object$entries.length; _i++) {
        var _Object$entries$_i = _slicedToArray(_Object$entries[_i], 2),
          key = _Object$entries$_i[0],
          value = _Object$entries$_i[1];
        payload.set(key, value);
      }
      console.debug("EditForm::ajax_submit --> ", payload);
      var init = {
        method: "POST",
        body: payload
      };
      return this.ajax_request(form, ajax_url, init);
    }

    /**
     * execute ajax request
     */
  }, {
    key: "ajax_request",
    value: function ajax_request(form, url, init) {
      var _this2 = this;
      // send ajax request to server
      this.loading(true);
      var request = new Request(url, init);
      return fetch(request).then(function (response) {
        if (!response.ok) {
          return Promise.reject(response);
        }
        return response.json();
      }).then(function (data) {
        console.debug("EditForm::ajax_request --> ", data);
        _this2.update_form(form, data);
        _this2.loading(false);
      })["catch"](function (error) {
        console.error(error);
        _this2.loading(false);
      });
    }

    /**
     * get the current view name from the URL
     */
  }, {
    key: "get_view_name",
    value: function get_view_name() {
      var segments = location.pathname.split("/");
      return segments.pop();
    }

    /**
     * Toggle element disable
     */
  }, {
    key: "toggle_disable",
    value: function toggle_disable(el, toggle) {
      if (el) {
        el.disabled = toggle;
      }
    }

    /**
     * Checks if the element is a textarea field
     */
  }, {
    key: "is_textarea",
    value: function is_textarea(el) {
      return el.tagName == "TEXTAREA";
    }

    /**
     * Checks if the elment is a select field
     */
  }, {
    key: "is_select",
    value: function is_select(el) {
      return el.tagName == "SELECT";
    }

    /**
     * Checks if the element is a multiple select field
     */
  }, {
    key: "is_multiple_select",
    value: function is_multiple_select(el) {
      return this.is_select(el) && el.hasAttribute("multiple");
    }

    /**
     * Checks if the element is an input field
     */
  }, {
    key: "is_input",
    value: function is_input(el) {
      return el.tagName === "INPUT";
    }

    /**
     * Checks if the element is an input[type='text'] field
     */
  }, {
    key: "is_text",
    value: function is_text(el) {
      return this.is_input(el) && el.type === "text";
    }

    /**
     * Checks if the element is a button field
     */
  }, {
    key: "is_button",
    value: function is_button(el) {
      return el.tagName === "BUTTON";
    }

    /**
     * Checks if the element is an input[type='button'] field
     */
  }, {
    key: "is_input_button",
    value: function is_input_button(el) {
      return this.is_input(el) && el.type === "button";
    }

    /**
     * Checks if the element is an input[type='checkbox'] field
     */
  }, {
    key: "is_checkbox",
    value: function is_checkbox(el) {
      return this.is_input(el) && el.type === "checkbox";
    }

    /**
     * Checks if the element is an input[type='radio'] field
     */
  }, {
    key: "is_radio",
    value: function is_radio(el) {
      return this.is_input(el) && el.type === "radio";
    }

    /**
     * Checks if the element is a SENAITE reference field (textarea)
     */
  }, {
    key: "is_reference",
    value: function is_reference(el) {
      if (!this.is_textarea(el)) {
        return false;
      }
      // NOTE: This class is only used if the field is not hidden.
      // Otherwise, it behaves like a normal textarea field.
      return el.classList.contains("queryselectwidget-value");
    }

    /**
     * Checks if the element is a table row element
     */
  }, {
    key: "is_table_row",
    value: function is_table_row(el) {
      return el.tagName === "TR";
    }
    /**
     * event handler for `mutated` event
     *
     * Queues the mutations and debounces their processing. ReactJS widgets
     * (e.g. the remarks widget) re-render on each keystroke, which would
     * otherwise trigger a server notification and a "Loading" flicker per key.
     */
  }, {
    key: "on_mutated",
    value: function on_mutated(event) {
      console.debug("EditForm::on_mutated");
      this.mutation_queue.push({
        form: event.detail.form,
        mutations: event.detail.mutations || []
      });
      if (this.mutation_timer) {
        clearTimeout(this.mutation_timer);
      }
      this.mutation_timer = setTimeout(this.flush_mutations, MUTATION_DEBOUNCE);
    }

    /**
     * process the queued DOM mutations (debounced)
     */
  }, {
    key: "flush_mutations",
    value: function flush_mutations() {
      this.mutation_timer = null;
      var queue = this.mutation_queue;
      this.mutation_queue = [];
      // reduce multiple mutations on the same node to one
      var seen = [];
      var _iterator19 = _createForOfIteratorHelper(queue),
        _step19;
      try {
        for (_iterator19.s(); !(_step19 = _iterator19.n()).done;) {
          var entry = _step19.value;
          var form = entry.form;
          var _iterator20 = _createForOfIteratorHelper(entry.mutations),
            _step20;
          try {
            for (_iterator20.s(); !(_step20 = _iterator20.n()).done;) {
              var mutation = _step20.value;
              if (seen.indexOf(mutation.target) > -1) {
                continue;
              }
              seen = seen.concat(mutation.target);
              this.handle_mutation(form, mutation);
            }
          } catch (err) {
            _iterator20.e(err);
          } finally {
            _iterator20.f();
          }
        }
      } catch (err) {
        _iterator19.e(err);
      } finally {
        _iterator19.f();
      }
    }

    /**
     * event handler for `modified` event
     */
  }, {
    key: "on_modified",
    value: function on_modified(event) {
      console.debug("EditForm::on_modified");
      var form = event.detail.form;
      var field = event.detail.field;
      this.notify(form, field, "modified");
    }

    /**
     * event handler for `submit` event
     */
  }, {
    key: "on_submit",
    value: function on_submit(event) {
      var _this3 = this;
      console.debug("EditForm::on_submit");
      event.preventDefault();
      var data = {};
      var form = event.currentTarget.closest("form");
      // NOTE: submit input field not included in request form data!
      var submitter = event.submitter;
      if (submitter) {
        data[submitter.name] = submitter.value;
        // disable submit button during ajax call
        this.toggle_disable(submitter, true);
      }
      this.ajax_submit(form, data, "submit").then(function (response) {
        return (
          // enable submit button after ajax call again
          _this3.toggle_disable(submitter, false)
        );
      });
    }

    /**
     * event handler for `blur` event
     */
  }, {
    key: "on_blur",
    value: function on_blur(event) {
      console.debug("EditForm::on_blur");
      var el = event.currentTarget;
      this.modified(el);
    }

    /**
     * event handler for `click` event
     */
  }, {
    key: "on_click",
    value: function on_click(event) {
      console.debug("EditForm::on_click");
      var el = event.currentTarget;
      this.modified(el);
    }

    /**
     * event handler for `change` event
     */
  }, {
    key: "on_change",
    value: function on_change(event) {
      console.debug("EditForm::on_change");
      var el = event.currentTarget;
      this.modified(el);
    }

    /**
     * event handler for `select` event
     */
  }, {
    key: "on_reference_select",
    value: function on_reference_select(event) {
      console.debug("EditForm::on_reference_select");
      var el = event.currentTarget;
      // add the selected value to the list
      var selected = el.value.split("\n");
      selected = selected.concat(event.detail.value);
      el.value = selected.join("\n");
      this.modified(el);
    }

    /**
     * event handler for `deselect` event
     */
  }, {
    key: "on_reference_deselect",
    value: function on_reference_deselect(event) {
      console.debug("EditForm::on_reference_deselect");
      var el = event.currentTarget;
      // remove the delelected value from the list
      var selected = el.value.split("\n");
      var index = selected.indexOf(event.detail.value);
      if (index > -1) {
        selected.splice(index, 1);
      }
      el.value = selected.join("\n");
      this.modified(el);
    }
  }]);
}();
/* harmony default export */ const editform = (EditForm);
;// ./components/site.js
/* provided dependency */ var $ = __webpack_require__(669);
/* Please use this command to compile this file into the parent `js` directory:
    coffee --no-header -w -o ../ -c site.coffee
 */
var Site,
  bind = function bind(fn, me) {
    return function () {
      return fn.apply(me, arguments);
    };
  };
Site = function () {
  /**
   * Creates a new instance of Site
   */
  function Site() {
    this.set_cookie = bind(this.set_cookie, this);
    this.read_cookie = bind(this.read_cookie, this);
    this.authenticator = bind(this.authenticator, this);
    console.debug("Site::init");
  }

  /**
   * Returns the authenticator value
   */

  Site.prototype.authenticator = function () {
    var auth, url_params;
    auth = $("input[name='_authenticator']").val();
    if (!auth) {
      url_params = new URLSearchParams(window.location.search);
      auth = url_params.get("_authenticator");
    }
    return auth;
  };

  /**
   * Reads a cookie value
   * @param {name} the name of the cookie
   */

  Site.prototype.read_cookie = function (name) {
    var c, ca, i;
    console.debug("Site::read_cookie:" + name);
    name = name + '=';
    ca = document.cookie.split(';');
    i = 0;
    while (i < ca.length) {
      c = ca[i];
      while (c.charAt(0) === ' ') {
        c = c.substring(1);
      }
      if (c.indexOf(name) === 0) {
        return c.substring(name.length, c.length);
      }
      i++;
    }
    return null;
  };

  /**
   * Sets a cookie value
   * @param {name} the name of the cookie
   * @param {value} the value of the cookie
   */

  Site.prototype.set_cookie = function (name, value) {
    var d, expires;
    console.debug("Site::set_cookie:name=" + name + ", value=" + value);
    d = new Date();
    d.setTime(d.getTime() + 1 * 24 * 60 * 60 * 1000);
    expires = 'expires=' + d.toUTCString();
    document.cookie = name + '=' + value + ';' + expires + ';path=/';
  };

  /**
   * Add a notification message
   * @param {title} the title of the notification
   * @param {message} the message of the notification
   */

  Site.prototype.add_notification = function (title, message, options) {
    var el, parent, wrapper;
    if (options == null) {
      options = {};
    }
    options = Object.assign({
      animation: true,
      autohide: true,
      delay: 5000
    }, options);
    el = document.createElement("div");
    el.innerHTML = "<div class='toast bg-white' style='width:300px' role='alert' data-animation='" + options.animation + "' data-autohide='" + options.autohide + "' data-delay='" + options.delay + "'> <div class='toast-header'> <strong class='mr-auto'>" + title + "</strong> <button type='button' class='ml-2 mb-1 close' data-dismiss='toast' aria-label='Close'> <span aria-hidden='true'>&times;</span> </button> </div> <div class='toast-body'> " + _t(message) + " </div> </div>";
    el = el.firstElementChild;
    parent = document.querySelector(".toast-container");
    if (!parent) {
      parent = document.createElement("div");
      parent.innerHTML = "<div style='position: fixed; top: 0px; right: 0px; width=100%; z-index:100'> <div class='toast-container' style='position: absolute; top: 10px; right: 10px;'> </div> </div>";
      wrapper = document.querySelector(".container-fluid");
      wrapper.appendChild(parent);
      parent = parent.querySelector(".toast-container");
    }
    parent.appendChild(el);
    return $(el).toast("show");
  };
  return Site;
}();
/* harmony default export */ const site = (Site);
;// ./components/calculationeditform.js
function calculationeditform_typeof(o) { "@babel/helpers - typeof"; return calculationeditform_typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, calculationeditform_typeof(o); }
function calculationeditform_classCallCheck(a, n) { if (!(a instanceof n)) throw new TypeError("Cannot call a class as a function"); }
function calculationeditform_defineProperties(e, r) { for (var t = 0; t < r.length; t++) { var o = r[t]; o.enumerable = o.enumerable || !1, o.configurable = !0, "value" in o && (o.writable = !0), Object.defineProperty(e, calculationeditform_toPropertyKey(o.key), o); } }
function calculationeditform_createClass(e, r, t) { return r && calculationeditform_defineProperties(e.prototype, r), t && calculationeditform_defineProperties(e, t), Object.defineProperty(e, "prototype", { writable: !1 }), e; }
function calculationeditform_toPropertyKey(t) { var i = calculationeditform_toPrimitive(t, "string"); return "symbol" == calculationeditform_typeof(i) ? i : i + ""; }
function calculationeditform_toPrimitive(t, r) { if ("object" != calculationeditform_typeof(t) || !t) return t; var e = t[Symbol.toPrimitive]; if (void 0 !== e) { var i = e.call(t, r || "default"); if ("object" != calculationeditform_typeof(i)) return i; throw new TypeError("@@toPrimitive must return a primitive value."); } return ("string" === r ? String : Number)(t); }

var CalculationEditForm = /*#__PURE__*/function () {
  function CalculationEditForm() {
    calculationeditform_classCallCheck(this, CalculationEditForm);
    this.DataGrid = null;
    this.testParamTable = null;
    this.rawTestValue = null;
    this.rawTestInput = document.getElementById("form-widgets-raw_test_keywords");
    if (this.rawTestInput) {
      this.load();
    }
  }
  return calculationeditform_createClass(CalculationEditForm, [{
    key: "load",
    value: function load() {
      this.rawTestValue = this.rawTestInput.value;
      this.makeReadonlyTestKeywords();
      this.hideAAField();
      this.wrapRawTestInput(this);
    }
  }, {
    key: "getDataGridWidget",
    value: function getDataGridWidget() {
      if (!this.DataGrid) {
        this.DataGrid = window.widgets.datagrid;
      }
      return this.DataGrid;
    }
  }, {
    key: "getTestParamTable",
    value: function getTestParamTable() {
      if (!this.testParamTable) {
        this.testParamTable = external_jQuery_default()("tbody[data-name_prefix='form.widgets.test_parameters']")[0];
      }
      return this.testParamTable;
    }
  }, {
    key: "makeReadonlyTestKeywords",
    value: function makeReadonlyTestKeywords() {
      external_jQuery_default()("input[id^='form-widgets-test_parameters-']").filter("[id$='-widgets-keyword']").each(function (i, e) {
        return external_jQuery_default()(e).attr("readonly", true);
      });
      external_jQuery_default()("#form-widgets-test_result").attr("readonly", true);
    }
  }, {
    key: "hideAAField",
    value: function hideAAField() {
      external_jQuery_default()("tbody[data-name_prefix='form.widgets.test_parameters'] > tr").each(function (i, e) {
        if (!["AA", "TT"].includes(external_jQuery_default()(e).attr("data-index"))) {
          external_jQuery_default()(e).show();
        } else {
          external_jQuery_default()(e).hide();
        }
      });
    }

    /**
     * Overrides the native setter for the raw test input field to control the test parameter DataGrid.
     *
     * XXX: This feels a bit hacky and is dependent on how the editform.js sets the input value!
     * E.g. if we use therer the `native_set_value` method, this will not work.
     *
     * Maybe it would be better to react on the `input` event or do this in a mutation observer
     */
  }, {
    key: "wrapRawTestInput",
    value: function wrapRawTestInput(parent) {
      var nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
      Object.defineProperty(this.rawTestInput, "value", {
        set: function set(newValue) {
          // prevent maximum call stack size exceeded error by using the native setter
          nativeSetter.call(this, newValue);
          parent.rawTestValue = newValue;
          var keywords = newValue.split(",").filter(function (k) {
            return k;
          });
          var table = parent.getTestParamTable();
          var visibleRows = parent.getDataGridWidget().get_visible_rows(table);
          if (keywords.length === 0) {
            for (var i = 0; i < visibleRows.length - 1; i++) {
              parent.getDataGridWidget().remove_row(visibleRows[i]);
            }
          } else if (keywords.length > visibleRows.length - 1) {
            var newRows = keywords.length - visibleRows.length + 1;
            for (var _i = 0; _i < newRows; _i++) {
              parent.getDataGridWidget().auto_append_row(table);
            }
          } else if (keywords.length < visibleRows.length - 1) {
            for (var _i2 = 0; _i2 < visibleRows.length - 1; _i2++) {
              var row = external_jQuery_default()(visibleRows[_i2]).find("input[id$='-widgets-keyword']");
              if (row) {
                if (!keywords.includes(row === null || row === void 0 ? void 0 : row.val())) {
                  parent.getDataGridWidget().remove_row(visibleRows[_i2]);
                }
              }
            }
          }
          parent.hideAAField();
          parent.getDataGridWidget().trigger_custom_event("update_test_parameters", keywords);
          parent.makeReadonlyTestKeywords();
        },
        get: function get() {
          return parent.rawTestValue;
        }
      });
    }
  }]);
}();
/* harmony default export */ const calculationeditform = (CalculationEditForm);
;// external "React"
const external_React_namespaceObject = React;
var external_React_default = /*#__PURE__*/__webpack_require__.n(external_React_namespaceObject);
;// external "ReactDOM"
const external_ReactDOM_namespaceObject = ReactDOM;
;// ./sidebar/hooks/useSidebarState.js
function useSidebarState_slicedToArray(r, e) { return useSidebarState_arrayWithHoles(r) || useSidebarState_iterableToArrayLimit(r, e) || useSidebarState_unsupportedIterableToArray(r, e) || useSidebarState_nonIterableRest(); }
function useSidebarState_nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function useSidebarState_unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return useSidebarState_arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? useSidebarState_arrayLikeToArray(r, a) : void 0; } }
function useSidebarState_arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function useSidebarState_iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t["return"] && (u = t["return"](), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function useSidebarState_arrayWithHoles(r) { if (Array.isArray(r)) return r; }


/**
 * Custom hook for managing sidebar toggle state
 * Handles cookie persistence and keyboard navigation
 */
var useSidebarState = function useSidebarState() {
  var cookieKey = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : "sidebar-toggle";
  var _useState = (0,external_React_namespaceObject.useState)(function () {
      if (typeof window !== "undefined" && window.site) {
        return window.site.read_cookie(cookieKey) === "true";
      }
      return false;
    }),
    _useState2 = useSidebarState_slicedToArray(_useState, 2),
    isToggled = _useState2[0],
    setIsToggled = _useState2[1];
  var _useState3 = (0,external_React_namespaceObject.useState)(!isToggled),
    _useState4 = useSidebarState_slicedToArray(_useState3, 2),
    isMinimized = _useState4[0],
    setIsMinimized = _useState4[1];
  var toggle = (0,external_React_namespaceObject.useCallback)(function (value) {
    var newValue = typeof value === "boolean" ? value : !isToggled;
    setIsToggled(newValue);
    setIsMinimized(!newValue);
    if (typeof window !== "undefined" && window.site) {
      window.site.set_cookie(cookieKey, newValue);
    }
  }, [isToggled, cookieKey]);
  (0,external_React_namespaceObject.useEffect)(function () {
    var handleKeyDown = function handleKeyDown(event) {
      if ((event.ctrlKey || event.metaKey) && event.key === "b") {
        event.preventDefault();
        toggle();
      }
    };
    document.addEventListener("keydown", handleKeyDown);
    return function () {
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [toggle]);
  return {
    isToggled: isToggled,
    isMinimized: isMinimized,
    toggle: toggle
  };
};
;// ./sidebar/hooks/useSidebarResize.js
function useSidebarResize_slicedToArray(r, e) { return useSidebarResize_arrayWithHoles(r) || useSidebarResize_iterableToArrayLimit(r, e) || useSidebarResize_unsupportedIterableToArray(r, e) || useSidebarResize_nonIterableRest(); }
function useSidebarResize_nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function useSidebarResize_unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return useSidebarResize_arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? useSidebarResize_arrayLikeToArray(r, a) : void 0; } }
function useSidebarResize_arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function useSidebarResize_iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t["return"] && (u = t["return"](), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function useSidebarResize_arrayWithHoles(r) { if (Array.isArray(r)) return r; }


/**
 * Custom hook for managing sidebar resize functionality
 * Handles mouse drag events and width persistence via cookies
 */
var useSidebarResize = function useSidebarResize() {
  var minWidth = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : 200;
  var maxWidth = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : 600;
  var widthCookieKey = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : "sidebar-width";
  var _useState = (0,external_React_namespaceObject.useState)(function () {
      if (typeof window !== "undefined" && window.site) {
        var savedWidth = window.site.read_cookie(widthCookieKey);
        if (savedWidth) {
          var parsedWidth = parseInt(savedWidth, 10);
          if (parsedWidth >= minWidth && parsedWidth <= maxWidth) {
            return parsedWidth;
          }
        }
      }
      return 200;
    }),
    _useState2 = useSidebarResize_slicedToArray(_useState, 2),
    width = _useState2[0],
    setWidth = _useState2[1];
  var _useState3 = (0,external_React_namespaceObject.useState)(false),
    _useState4 = useSidebarResize_slicedToArray(_useState3, 2),
    isResizing = _useState4[0],
    setIsResizing = _useState4[1];
  var resizeStartX = (0,external_React_namespaceObject.useRef)(0);
  var resizeStartWidth = (0,external_React_namespaceObject.useRef)(0);
  var startResize = (0,external_React_namespaceObject.useCallback)(function (event) {
    event.preventDefault();
    setIsResizing(true);
    resizeStartX.current = event.clientX;
    resizeStartWidth.current = width;
    document.body.style.userSelect = "none";
    document.body.style.cursor = "ew-resize";
  }, [width]);
  var resize = (0,external_React_namespaceObject.useCallback)(function (event) {
    if (!isResizing) return;
    event.preventDefault();
    var delta = event.clientX - resizeStartX.current;
    var newWidth = resizeStartWidth.current + delta;
    newWidth = Math.max(minWidth, Math.min(maxWidth, newWidth));
    setWidth(newWidth);
  }, [isResizing, minWidth, maxWidth]);
  var stopResize = (0,external_React_namespaceObject.useCallback)(function () {
    if (!isResizing) return;
    setIsResizing(false);
    document.body.style.userSelect = "";
    document.body.style.cursor = "";
    if (typeof window !== "undefined" && window.site) {
      window.site.set_cookie(widthCookieKey, width);
    }
  }, [isResizing, width, widthCookieKey]);
  (0,external_React_namespaceObject.useEffect)(function () {
    if (isResizing) {
      document.addEventListener("mousemove", resize);
      document.addEventListener("mouseup", stopResize);
      return function () {
        document.removeEventListener("mousemove", resize);
        document.removeEventListener("mouseup", stopResize);
      };
    }
  }, [isResizing, resize, stopResize]);
  return {
    width: width,
    isResizing: isResizing,
    startResize: startResize
  };
};
;// ./sidebar/hooks/useNavigation.js
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i["return"]) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
function useNavigation_slicedToArray(r, e) { return useNavigation_arrayWithHoles(r) || useNavigation_iterableToArrayLimit(r, e) || useNavigation_unsupportedIterableToArray(r, e) || useNavigation_nonIterableRest(); }
function useNavigation_nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function useNavigation_unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return useNavigation_arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? useNavigation_arrayLikeToArray(r, a) : void 0; } }
function useNavigation_arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function useNavigation_iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t["return"] && (u = t["return"](), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function useNavigation_arrayWithHoles(r) { if (Array.isArray(r)) return r; }


/**
 * Custom hook for fetching and managing navigation data
 */
var useNavigation = function useNavigation() {
  var _useState = (0,external_React_namespaceObject.useState)([]),
    _useState2 = useNavigation_slicedToArray(_useState, 2),
    navigationData = _useState2[0],
    setNavigationData = _useState2[1];
  var _useState3 = (0,external_React_namespaceObject.useState)(false),
    _useState4 = useNavigation_slicedToArray(_useState3, 2),
    isLoading = _useState4[0],
    setIsLoading = _useState4[1];
  var _useState5 = (0,external_React_namespaceObject.useState)(null),
    _useState6 = useNavigation_slicedToArray(_useState5, 2),
    error = _useState6[0],
    setError = _useState6[1];
  var fetchNavigation = (0,external_React_namespaceObject.useCallback)(/*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee() {
    var showMore,
      portalUrl,
      currentUrl,
      url,
      response,
      result,
      _args = arguments,
      _t;
    return _regenerator().w(function (_context) {
      while (1) switch (_context.p = _context.n) {
        case 0:
          showMore = _args.length > 0 && _args[0] !== undefined ? _args[0] : false;
          _context.p = 1;
          setIsLoading(true);
          setError(null);
          portalUrl = window.portal_url || "/";
          currentUrl = document.body.dataset.baseUrl || window.location.href;
          url = "".concat(portalUrl, "/@@sidebar-navigation-json?current_url=").concat(encodeURIComponent(currentUrl), "&show_more=").concat(showMore);
          _context.n = 2;
          return fetch(url, {
            method: "GET",
            headers: {
              "Content-Type": "application/json"
            },
            credentials: "same-origin"
          });
        case 2:
          response = _context.v;
          if (response.ok) {
            _context.n = 3;
            break;
          }
          throw new Error("HTTP error! status: ".concat(response.status));
        case 3:
          _context.n = 4;
          return response.json();
        case 4:
          result = _context.v;
          if (result.success) {
            _context.n = 5;
            break;
          }
          throw new Error(result.error || "Failed to fetch navigation");
        case 5:
          setNavigationData(result.data);
          _context.n = 7;
          break;
        case 6:
          _context.p = 6;
          _t = _context.v;
          console.error("Error fetching sidebar navigation:", _t);
          setError(_t.message);
        case 7:
          _context.p = 7;
          setIsLoading(false);
          return _context.f(7);
        case 8:
          return _context.a(2);
      }
    }, _callee, null, [[1, 6, 7, 8]]);
  })), []);
  (0,external_React_namespaceObject.useEffect)(function () {
    fetchNavigation();
  }, [fetchNavigation]);
  var showMoreItems = (0,external_React_namespaceObject.useCallback)(function () {
    fetchNavigation(true);
  }, [fetchNavigation]);
  return {
    navigationData: navigationData,
    isLoading: isLoading,
    error: error,
    showMoreItems: showMoreItems
  };
};
;// ./sidebar/components/SidebarHeader.js


/**
 * Sidebar Header Component
 * Displays the toggle button for expanding/collapsing the sidebar
 */
var SidebarHeader = function SidebarHeader(_ref) {
  var isToggled = _ref.isToggled,
    onToggle = _ref.onToggle;
  var iconClass = isToggled ? "fa fa-times" : "fa fa-bars";
  return /*#__PURE__*/external_React_default().createElement("div", {
    id: "sidebar-header"
  }, /*#__PURE__*/external_React_default().createElement("button", {
    type: "button",
    onClick: onToggle,
    title: "Toggle sidebar (Ctrl/Cmd+B)"
  }, /*#__PURE__*/external_React_default().createElement("i", {
    className: "sidebar-toggle-icon ".concat(iconClass)
  })));
};
;// ./sidebar/components/SidebarSearch.js
function SidebarSearch_slicedToArray(r, e) { return SidebarSearch_arrayWithHoles(r) || SidebarSearch_iterableToArrayLimit(r, e) || SidebarSearch_unsupportedIterableToArray(r, e) || SidebarSearch_nonIterableRest(); }
function SidebarSearch_nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function SidebarSearch_unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return SidebarSearch_arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? SidebarSearch_arrayLikeToArray(r, a) : void 0; } }
function SidebarSearch_arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function SidebarSearch_iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t["return"] && (u = t["return"](), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function SidebarSearch_arrayWithHoles(r) { if (Array.isArray(r)) return r; }


/**
 * Sidebar Search Component
 * Provides search functionality for filtering navigation items
 */
var SidebarSearch = function SidebarSearch(_ref) {
  var onSearch = _ref.onSearch,
    onFocus = _ref.onFocus,
    onBlur = _ref.onBlur;
  var _useState = (0,external_React_namespaceObject.useState)(""),
    _useState2 = SidebarSearch_slicedToArray(_useState, 2),
    query = _useState2[0],
    setQuery = _useState2[1];
  var handleInput = (0,external_React_namespaceObject.useCallback)(function (event) {
    var value = event.target.value;
    setQuery(value);
    if (onSearch) {
      onSearch(value);
    }
  }, [onSearch]);
  var handleFocus = (0,external_React_namespaceObject.useCallback)(function (event) {
    if (onFocus) {
      onFocus(event);
    }
  }, [onFocus]);
  var handleBlur = (0,external_React_namespaceObject.useCallback)(function (event) {
    if (onBlur) {
      onBlur(event);
    }
  }, [onBlur]);
  return /*#__PURE__*/external_React_default().createElement("div", {
    id: "sidebar-search-container"
  }, /*#__PURE__*/external_React_default().createElement("input", {
    type: "text",
    id: "sidebar-search",
    placeholder: window._t ? window._t("Search...") : "Search...",
    value: query,
    onInput: handleInput,
    onFocus: handleFocus,
    onBlur: handleBlur
  }));
};
;// ./sidebar/components/SidebarItem.js
function SidebarItem_slicedToArray(r, e) { return SidebarItem_arrayWithHoles(r) || SidebarItem_iterableToArrayLimit(r, e) || SidebarItem_unsupportedIterableToArray(r, e) || SidebarItem_nonIterableRest(); }
function SidebarItem_nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function SidebarItem_unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return SidebarItem_arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? SidebarItem_arrayLikeToArray(r, a) : void 0; } }
function SidebarItem_arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function SidebarItem_iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t["return"] && (u = t["return"](), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function SidebarItem_arrayWithHoles(r) { if (Array.isArray(r)) return r; }


/**
 * Sidebar Item Component
 * Renders a single navigation item with optional children (recursive)
 */
var _SidebarItem = function SidebarItem(_ref) {
  var item = _ref.item,
    _ref$level = _ref.level,
    level = _ref$level === void 0 ? 1 : _ref$level,
    _ref$searchQuery = _ref.searchQuery,
    searchQuery = _ref$searchQuery === void 0 ? "" : _ref$searchQuery,
    onShowMore = _ref.onShowMore;
  var _useState = (0,external_React_namespaceObject.useState)(item.is_current || item.is_parent),
    _useState2 = SidebarItem_slicedToArray(_useState, 2),
    isExpanded = _useState2[0],
    setIsExpanded = _useState2[1];
  var hasChildren = item.children && item.children.length > 0;
  var toggleExpanded = (0,external_React_namespaceObject.useCallback)(function (event) {
    var target = event.target;
    var isCaret = target.classList.contains("caret") || target.closest(".caret");
    if (isCaret) {
      event.preventDefault();
      setIsExpanded(!isExpanded);
    }
  }, [isExpanded]);
  var itemClasses = ["navTreeItem"];
  if (item.is_current) {
    itemClasses.push("active", "navTreeCurrentNode");
  }
  if (item.is_parent) {
    itemClasses.push("navTreeCurrentParent");
  }
  if (hasChildren) {
    itemClasses.push("navTreeFolderish");
    itemClasses.push(isExpanded ? "expanded" : "collapsed");
  }
  var shouldDisplay = !searchQuery || item.title.toLowerCase().includes(searchQuery.toLowerCase());
  if (!shouldDisplay && !hasChildren) {
    return null;
  }
  var portalUrl = window.portal_url || "";
  return /*#__PURE__*/external_React_default().createElement("li", {
    className: itemClasses.join(" ")
  }, /*#__PURE__*/external_React_default().createElement("a", {
    href: item.url,
    className: "navTreeLink",
    "data-id": item.id,
    "data-portal-type": item.portal_type,
    title: item.description || "",
    onClick: hasChildren ? toggleExpanded : undefined
  }, item.icon && /*#__PURE__*/external_React_default().createElement("span", {
    className: "node-icon"
  }, /*#__PURE__*/external_React_default().createElement("img", {
    src: "".concat(portalUrl, "/").concat(item.icon),
    alt: "",
    className: "nav-icon"
  })), /*#__PURE__*/external_React_default().createElement("span", {
    className: level === 1 ? "node-title" : "child-title"
  }, item.title), hasChildren && /*#__PURE__*/external_React_default().createElement("span", {
    className: "caret"
  })), hasChildren && /*#__PURE__*/external_React_default().createElement("ul", {
    className: "nav-level-".concat(level + 1)
  }, item.children.map(function (child) {
    return /*#__PURE__*/external_React_default().createElement(_SidebarItem, {
      key: child.id,
      item: child,
      level: level + 1,
      searchQuery: searchQuery,
      onShowMore: onShowMore
    });
  }), item.has_more && /*#__PURE__*/external_React_default().createElement("li", {
    className: "navTreeItem load-more-item"
  }, /*#__PURE__*/external_React_default().createElement("a", {
    href: "#",
    className: "navTreeLink load-more-link",
    "data-parent-id": item.id,
    title: "Click to load all items in this folder",
    onClick: function onClick(event) {
      event.preventDefault();
      if (onShowMore) {
        onShowMore();
      }
    }
  }, "Show more..."))));
};

;// ./sidebar/components/SidebarNavigation.js



/**
 * Sidebar Navigation Component
 * Renders the navigation tree
 */
var SidebarNavigation = function SidebarNavigation(_ref) {
  var navigationData = _ref.navigationData,
    searchQuery = _ref.searchQuery,
    onShowMore = _ref.onShowMore;
  if (!navigationData || navigationData.length === 0) {
    return null;
  }
  return /*#__PURE__*/external_React_default().createElement("div", {
    className: "sidebar-navigation"
  }, /*#__PURE__*/external_React_default().createElement("ul", {
    className: "nav-level-1"
  }, navigationData.map(function (item) {
    return /*#__PURE__*/external_React_default().createElement(_SidebarItem, {
      key: item.id,
      item: item,
      level: 1,
      searchQuery: searchQuery,
      onShowMore: onShowMore
    });
  })));
};
;// ./sidebar/Sidebar.js
function Sidebar_slicedToArray(r, e) { return Sidebar_arrayWithHoles(r) || Sidebar_iterableToArrayLimit(r, e) || Sidebar_unsupportedIterableToArray(r, e) || Sidebar_nonIterableRest(); }
function Sidebar_nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function Sidebar_unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return Sidebar_arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? Sidebar_arrayLikeToArray(r, a) : void 0; } }
function Sidebar_arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function Sidebar_iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t["return"] && (u = t["return"](), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function Sidebar_arrayWithHoles(r) { if (Array.isArray(r)) return r; }









/**
 * Main Sidebar Component
 * Modern sidebar with smooth animations and collapsible sections
 *
 * Features:
 * - Toggle button for persistent state
 * - Smooth CSS transitions
 * - Keyboard navigation support (Ctrl/Cmd+B)
 * - Collapsible navigation sections
 * - Search functionality for navigation items
 * - Resizable sidebar with drag handle
 * - Mobile-responsive with overlay and slide-in animation
 * - Body scroll lock when mobile sidebar is open
 */
var Sidebar = function Sidebar() {
  var _useSidebarState = useSidebarState(),
    isToggled = _useSidebarState.isToggled,
    isMinimized = _useSidebarState.isMinimized,
    toggle = _useSidebarState.toggle;
  var _useSidebarResize = useSidebarResize(),
    width = _useSidebarResize.width,
    isResizing = _useSidebarResize.isResizing,
    startResize = _useSidebarResize.startResize;
  var _useNavigation = useNavigation(),
    navigationData = _useNavigation.navigationData,
    isLoading = _useNavigation.isLoading,
    error = _useNavigation.error,
    showMoreItems = _useNavigation.showMoreItems;
  var _useState = (0,external_React_namespaceObject.useState)(""),
    _useState2 = Sidebar_slicedToArray(_useState, 2),
    searchQuery = _useState2[0],
    setSearchQuery = _useState2[1];
  var _useState3 = (0,external_React_namespaceObject.useState)(false),
    _useState4 = Sidebar_slicedToArray(_useState3, 2),
    isSearchActive = _useState4[0],
    setIsSearchActive = _useState4[1];
  var _useState5 = (0,external_React_namespaceObject.useState)(false),
    _useState6 = Sidebar_slicedToArray(_useState5, 2),
    isMobile = _useState6[0],
    setIsMobile = _useState6[1];

  // Detect mobile viewport
  (0,external_React_namespaceObject.useEffect)(function () {
    var checkMobile = function checkMobile() {
      setIsMobile(window.innerWidth <= 768);
    };
    checkMobile();
    window.addEventListener("resize", checkMobile);
    return function () {
      return window.removeEventListener("resize", checkMobile);
    };
  }, []);

  // Apply classes and styles to the #sidebar container
  (0,external_React_namespaceObject.useEffect)(function () {
    var container = document.getElementById("sidebar");
    if (!container) return;
    var classes = [];
    if (isMinimized) classes.push("minimized");
    if (isToggled) classes.push("toggled");
    if (isLoading) classes.push("loading");
    if (isSearchActive) classes.push("search-active");
    container.className = classes.join(" ");
    container.style.width = isMinimized ? "50px" : "".concat(width, "px");
  }, [isMinimized, isToggled, isLoading, isSearchActive, width]);

  // Manage body scroll for mobile
  (0,external_React_namespaceObject.useEffect)(function () {
    if (isMobile && isToggled && !isMinimized) {
      document.body.style.overflow = "hidden";
    } else {
      document.body.style.overflow = "";
    }
    return function () {
      document.body.style.overflow = "";
    };
  }, [isMobile, isToggled, isMinimized]);

  // Handle backdrop click to close sidebar
  var handleBackdropClick = (0,external_React_namespaceObject.useCallback)(function () {
    if (isMobile) {
      toggle(false);
    }
  }, [isMobile, toggle]);
  var handleSearch = (0,external_React_namespaceObject.useCallback)(function (query) {
    setSearchQuery(query);
  }, []);
  var handleSearchFocus = (0,external_React_namespaceObject.useCallback)(function () {
    setIsSearchActive(true);
    // Expand sidebar if minimized
    if (isMinimized) {
      toggle(true);
    }
  }, [isMinimized, toggle]);
  var handleSearchBlur = (0,external_React_namespaceObject.useCallback)(function () {
    setTimeout(function () {
      setIsSearchActive(false);
    }, 200);
  }, []);
  var handleToggle = (0,external_React_namespaceObject.useCallback)(function () {
    toggle();
  }, [toggle]);
  var showBackdrop = isMobile && isToggled && !isMinimized;
  return /*#__PURE__*/external_React_default().createElement((external_React_default()).Fragment, null, /*#__PURE__*/(0,external_ReactDOM_namespaceObject.createPortal)(/*#__PURE__*/external_React_default().createElement("div", {
    className: "sidebar-backdrop ".concat(showBackdrop ? "show" : ""),
    onClick: handleBackdropClick
  }), document.body), isMobile && !isToggled ? (/*#__PURE__*/(0,external_ReactDOM_namespaceObject.createPortal)(/*#__PURE__*/external_React_default().createElement(SidebarHeader, {
    isToggled: isToggled,
    onToggle: handleToggle
  }), document.body)) : /*#__PURE__*/external_React_default().createElement(SidebarHeader, {
    isToggled: isToggled,
    onToggle: handleToggle
  }), /*#__PURE__*/external_React_default().createElement(SidebarSearch, {
    onSearch: handleSearch,
    onFocus: handleSearchFocus,
    onBlur: handleSearchBlur
  }), isLoading && /*#__PURE__*/external_React_default().createElement("div", {
    className: "sidebar-loading"
  }, /*#__PURE__*/external_React_default().createElement("div", {
    className: "spinner"
  }), /*#__PURE__*/external_React_default().createElement("span", null, window._t ? window._t("Loading navigation...") : "Loading navigation...")), error && /*#__PURE__*/external_React_default().createElement("div", {
    className: "sidebar-error"
  }, error), !isLoading && !error && /*#__PURE__*/external_React_default().createElement(SidebarNavigation, {
    navigationData: navigationData,
    searchQuery: searchQuery,
    onShowMore: showMoreItems
  }), !isMinimized && /*#__PURE__*/external_React_default().createElement("div", {
    className: "resize-handle ".concat(isResizing ? "resizing" : ""),
    onMouseDown: startResize
  }));
};
/* harmony default export */ const sidebar_Sidebar = (Sidebar);
;// ./sidebar/index.js
/* unused harmony import specifier */ var sidebar_Sidebar_0;
/**
 * SENAITE Sidebar - React Component
 *
 * Entry point for the sidebar React component
 */






/**
 * Sidebar API object for backwards compatibility
 * Provides access to the sidebar component and common operations
 */
var SidebarAPI = {
  component: null,
  container: null,
  /**
   * Get the sidebar DOM element
   */
  getElement: function getElement() {
    return document.getElementById("sidebar");
  },
  /**
   * Check if sidebar is toggled (expanded)
   */
  isToggled: function isToggled() {
    if (typeof window !== "undefined" && window.site) {
      return window.site.read_cookie("sidebar-toggle") === "true";
    }
    return false;
  },
  /**
   * Check if sidebar is minimized
   */
  isMinimized: function isMinimized() {
    var el = this.getElement();
    return el ? el.classList.contains("minimized") : true;
  },
  /**
   * Toggle sidebar programmatically
   */
  toggle: function toggle(value) {
    var el = this.getElement();
    if (!el) return;
    var newValue = typeof value === "boolean" ? value : !this.isToggled();
    if (typeof window !== "undefined" && window.site) {
      window.site.set_cookie("sidebar-toggle", newValue);
    }
    if (newValue) {
      el.classList.add("toggled");
      el.classList.remove("minimized");
    } else {
      el.classList.remove("toggled");
      el.classList.add("minimized");
    }
  },
  /**
   * Minimize sidebar
   */
  minimize: function minimize() {
    this.toggle(false);
  },
  /**
   * Maximize sidebar
   */
  maximize: function maximize() {
    this.toggle(true);
  }
};

/**
 * Initialize the sidebar React component
 */
var initSidebar = function initSidebar() {
  var sidebarContainer = document.getElementById("sidebar");
  if (!sidebarContainer) {
    console.warn("Sidebar container not found");
    return null;
  }
  SidebarAPI.container = sidebarContainer;

  // Use React 18 createRoot API
  var root = (0,external_ReactDOM_namespaceObject.createRoot)(sidebarContainer);
  root.render(/*#__PURE__*/external_React_default().createElement(sidebar_Sidebar, null));
  return SidebarAPI;
};
/* harmony default export */ const sidebar = ((/* unused pure expression or super */ null && (sidebar_Sidebar_0)));
;// ./workflow-menu/WorkflowMenu.js
function WorkflowMenu_slicedToArray(r, e) { return WorkflowMenu_arrayWithHoles(r) || WorkflowMenu_iterableToArrayLimit(r, e) || WorkflowMenu_unsupportedIterableToArray(r, e) || WorkflowMenu_nonIterableRest(); }
function WorkflowMenu_nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function WorkflowMenu_unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return WorkflowMenu_arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? WorkflowMenu_arrayLikeToArray(r, a) : void 0; } }
function WorkflowMenu_arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function WorkflowMenu_iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t["return"] && (u = t["return"](), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function WorkflowMenu_arrayWithHoles(r) { if (Array.isArray(r)) return r; }
/**
 * Workflow Menu
 *
 * Renders the workflow dropdown in the SENAITE content toolbar.
 *
 * The current review state is shown immediately from props (seeded by
 * the server-side `data-*` attributes). The transition list is fetched
 * lazily from `@@workflow_menu_data` the first time the menu opens and
 * cached per-UID for the lifetime of the page.
 */


var transitionCache = new Map();
var WorkflowMenu_t = function t(msg) {
  if (typeof window !== "undefined" && typeof window._t === "function") {
    return window._t(msg);
  }
  return msg;
};
var WorkflowMenu = function WorkflowMenu(props) {
  var uid = props.uid,
    stateTitle = props.stateTitle,
    stateClass = props.stateClass,
    fetchUrl = props.fetchUrl;
  var _useState = (0,external_React_namespaceObject.useState)(false),
    _useState2 = WorkflowMenu_slicedToArray(_useState, 2),
    open = _useState2[0],
    setOpen = _useState2[1];
  var _useState3 = (0,external_React_namespaceObject.useState)(transitionCache.has(uid) ? "ready" : "idle"),
    _useState4 = WorkflowMenu_slicedToArray(_useState3, 2),
    status = _useState4[0],
    setStatus = _useState4[1];
  var _useState5 = (0,external_React_namespaceObject.useState)(transitionCache.get(uid) || null),
    _useState6 = WorkflowMenu_slicedToArray(_useState5, 2),
    items = _useState6[0],
    setItems = _useState6[1];
  var _useState7 = (0,external_React_namespaceObject.useState)(null),
    _useState8 = WorkflowMenu_slicedToArray(_useState7, 2),
    error = _useState8[0],
    setError = _useState8[1];
  var wrapperRef = (0,external_React_namespaceObject.useRef)(null);
  var load = (0,external_React_namespaceObject.useCallback)(function () {
    if (items !== null || status === "loading") {
      return;
    }
    setStatus("loading");
    fetch(fetchUrl, {
      credentials: "same-origin",
      headers: {
        Accept: "application/json"
      }
    }).then(function (response) {
      if (!response.ok) {
        throw new Error("HTTP ".concat(response.status));
      }
      return response.json();
    }).then(function (data) {
      var list = data.transitions || [];
      transitionCache.set(uid, list);
      setItems(list);
      setStatus("ready");
    })["catch"](function (err) {
      setError(err.message || String(err));
      setStatus("error");
    });
  }, [uid, fetchUrl, items, status]);
  var toggle = (0,external_React_namespaceObject.useCallback)(function (event) {
    event.preventDefault();
    setOpen(function (prev) {
      var next = !prev;
      if (next) {
        load();
      }
      return next;
    });
  }, [load]);
  (0,external_React_namespaceObject.useEffect)(function () {
    if (!open) {
      return undefined;
    }
    var onClickAway = function onClickAway(event) {
      var node = wrapperRef.current;
      if (node && !node.contains(event.target)) {
        setOpen(false);
      }
    };
    var onKeyDown = function onKeyDown(event) {
      if (event.key === "Escape") {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", onClickAway);
    document.addEventListener("keydown", onKeyDown);
    return function () {
      document.removeEventListener("mousedown", onClickAway);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [open]);
  return /*#__PURE__*/external_React_default().createElement("div", {
    ref: wrapperRef,
    className: "d-inline-flex"
  }, /*#__PURE__*/external_React_default().createElement("a", {
    href: "#",
    className: "nav-link dropdown-toggle workflow-menu-toggle",
    role: "button",
    "aria-haspopup": "true",
    "aria-expanded": open,
    onClick: toggle
  }, /*#__PURE__*/external_React_default().createElement("span", {
    className: "".concat(stateClass || "", " tb-state")
  }, stateTitle)), /*#__PURE__*/external_React_default().createElement("div", {
    className: "dropdown-menu workflow-menu-dropdown".concat(open ? " show" : "")
  }, status === "loading" && /*#__PURE__*/external_React_default().createElement("span", {
    className: "dropdown-item-text text-muted"
  }, /*#__PURE__*/external_React_default().createElement("i", {
    className: "fas fa-spinner fa-spin mr-1"
  }), WorkflowMenu_t("Loading…")), status === "error" && /*#__PURE__*/external_React_default().createElement("span", {
    className: "dropdown-item-text text-danger"
  }, error), status === "ready" && items && items.length === 0 && /*#__PURE__*/external_React_default().createElement("span", {
    className: "dropdown-item-text text-muted"
  }, WorkflowMenu_t("No transitions available")), status === "ready" && items && items.map(function (transition) {
    return /*#__PURE__*/external_React_default().createElement("a", {
      key: transition.id,
      className: "dropdown-item",
      href: transition.url,
      title: transition.description
    }, transition.title);
  })));
};
/* harmony default export */ const workflow_menu_WorkflowMenu = (WorkflowMenu);
;// ./workflow-menu/index.js
/* unused harmony import specifier */ var workflow_menu_WorkflowMenu_0;
/**
 * SENAITE Workflow Menu - React Component bootstrapper
 *
 * Mounts a `WorkflowMenu` component on every `.senaite-workflow-menu`
 * placeholder rendered by the `plone.contentmenu` viewlet. The
 * placeholder carries the current review state via data-* attributes so
 * the menu is visible immediately; the list of allowed transitions is
 * fetched lazily from `@@workflow_menu_data` on first open.
 */





var initWorkflowMenus = function initWorkflowMenus() {
  var nodes = document.querySelectorAll(".senaite-workflow-menu");
  var roots = [];
  nodes.forEach(function (node) {
    var props = {
      uid: node.dataset.uid,
      stateTitle: node.dataset.stateTitle || "",
      stateClass: node.dataset.stateClass || "",
      fetchUrl: node.dataset.fetchUrl
    };
    // Replace the SSR skeleton so React owns the whole subtree.
    node.replaceChildren();
    var root = (0,external_ReactDOM_namespaceObject.createRoot)(node);
    root.render(/*#__PURE__*/external_React_default().createElement(workflow_menu_WorkflowMenu, props));
    roots.push(root);
  });
  return roots;
};
/* harmony default export */ const workflow_menu = ((/* unused pure expression or super */ null && (workflow_menu_WorkflowMenu_0)));
;// ./components/formtabbing.js
function formtabbing_typeof(o) { "@babel/helpers - typeof"; return formtabbing_typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, formtabbing_typeof(o); }
function formtabbing_classCallCheck(a, n) { if (!(a instanceof n)) throw new TypeError("Cannot call a class as a function"); }
function formtabbing_defineProperties(e, r) { for (var t = 0; t < r.length; t++) { var o = r[t]; o.enumerable = o.enumerable || !1, o.configurable = !0, "value" in o && (o.writable = !0), Object.defineProperty(e, formtabbing_toPropertyKey(o.key), o); } }
function formtabbing_createClass(e, r, t) { return r && formtabbing_defineProperties(e.prototype, r), t && formtabbing_defineProperties(e, t), Object.defineProperty(e, "prototype", { writable: !1 }), e; }
function formtabbing_toPropertyKey(t) { var i = formtabbing_toPrimitive(t, "string"); return "symbol" == formtabbing_typeof(i) ? i : i + ""; }
function formtabbing_toPrimitive(t, r) { if ("object" != formtabbing_typeof(t) || !t) return t; var e = t[Symbol.toPrimitive]; if (void 0 !== e) { var i = e.call(t, r || "default"); if ("object" != formtabbing_typeof(i)) return i; throw new TypeError("@@toPrimitive must return a primitive value."); } return ("string" === r ? String : Number)(t); }
/**
 * Form Tabbing Component
 *
 * Remembers the active tab in Bootstrap 4 tabs for content edit and view forms
 * Uses URL parameter to persist the active tab across page reloads
 */


var FormTabbing = /*#__PURE__*/function () {
  function FormTabbing() {
    formtabbing_classCallCheck(this, FormTabbing);
    this.TAB_PARAM = "tab";
  }

  /**
   * Initialize form tabbing functionality
   */
  return formtabbing_createClass(FormTabbing, [{
    key: "init",
    value: function init() {
      var _this = this;
      var tabs = document.querySelectorAll(".nav-tabs a[data-toggle='tab']");
      if (tabs.length === 0) {
        console.debug("FormTabbing: No tabs found");
        return;
      }
      console.debug("FormTabbing: Found ".concat(tabs.length, " tabs"));

      // Restore the tab from URL parameter
      this.restoreActiveTab();

      // Update tab links to modify URL on click
      this.updateTabLinks();

      // Add tab parameter to edit/view links
      this.addTabParameterToLinks();

      // Listen for tab changes to update URL and links using jQuery
      external_jQuery_default()(".nav-tabs a[data-toggle='tab']").on("shown.bs.tab", function (event) {
        var tabId = external_jQuery_default()(event.target).attr("id");
        if (tabId) {
          _this.updateUrl(tabId);
          // Update edit/view links when tab changes
          _this.addTabParameterToLinks();
        }
      });
      console.debug("FormTabbing: Tab memory initialized");
    }

    /**
     * Get URL parameter value
     */
  }, {
    key: "getUrlParameter",
    value: function getUrlParameter(name) {
      var url = window.location.href;
      var escapedName = name.replace(/[\[\]]/g, "\\$&");
      var regex = new RegExp("[?&]" + escapedName + "(=([^&#]*)|&|#|$)");
      var results = regex.exec(url);
      if (!results) return null;
      if (!results[2]) return "";
      return decodeURIComponent(results[2].replace(/\+/g, " "));
    }

    /**
     * Update URL parameter
     */
  }, {
    key: "updateUrlParameter",
    value: function updateUrlParameter(param, value) {
      var url = window.location.href;
      var regex = new RegExp("([?&])" + param + "=.*?(&|$)", "i");
      var separator = url.indexOf("?") !== -1 ? "&" : "?";
      if (url.match(regex)) {
        return url.replace(regex, "$1" + param + "=" + value + "$2");
      } else {
        return url + separator + param + "=" + value;
      }
    }

    /**
     * Update URL with tab parameter
     */
  }, {
    key: "updateUrl",
    value: function updateUrl(tabId) {
      try {
        var newUrl = this.updateUrlParameter(this.TAB_PARAM, tabId);
        window.history.replaceState({}, "", newUrl);
        console.debug("FormTabbing: Updated URL with tab ".concat(tabId));
      } catch (e) {
        console.warn("FormTabbing: Could not update URL", e);
      }
    }

    /**
     * Restore the active tab from URL parameter
     */
  }, {
    key: "restoreActiveTab",
    value: function restoreActiveTab() {
      try {
        var savedTabId = this.getUrlParameter(this.TAB_PARAM);
        if (!savedTabId) {
          console.debug("FormTabbing: No tab parameter found");
          return;
        }
        console.debug("FormTabbing: Looking for tab with ID: ".concat(savedTabId));
        var savedTab = document.getElementById(savedTabId);
        if (!savedTab) {
          console.warn("FormTabbing: Tab element with ID \"".concat(savedTabId, "\" not found in DOM"));

          // Try to find the tab by href as fallback
          var tabByHref = document.querySelector(".nav-tabs a[href=\"#".concat(savedTabId.replace("-tab", ""), "\"]"));
          if (tabByHref) {
            savedTab = tabByHref;
            console.debug("FormTabbing: Found tab by href instead");
          } else {
            console.warn("FormTabbing: Could not find tab by href either");
            return;
          }
        }

        // Remove active class from all tabs and panes
        external_jQuery_default()(".nav-tabs .nav-link").removeClass("active");
        external_jQuery_default()(".tab-pane").removeClass("active show");

        // Activate the saved tab using Bootstrap's tab method
        external_jQuery_default()(savedTab).tab("show");
        console.info("FormTabbing: Successfully restored tab ".concat(savedTabId));
      } catch (e) {
        console.error("FormTabbing: Error restoring tab", e);
      }
    }

    /**
     * Update tab links to modify URL on click
     */
  }, {
    key: "updateTabLinks",
    value: function updateTabLinks() {
      var _this2 = this;
      external_jQuery_default()(".nav-tabs a[data-toggle='tab']").on("click", function (event) {
        var $tab = external_jQuery_default()(event.currentTarget);
        var originalHref = $tab.attr("href");
        if (originalHref && originalHref.indexOf("#") === 0) {
          var tabId = $tab.attr("id");
          if (tabId) {
            // Update URL with the new tab parameter
            setTimeout(function () {
              _this2.updateUrl(tabId);
            }, 0);
          }
        }
      });
    }

    /**
     * Add tab parameter to edit/view links
     */
  }, {
    key: "addTabParameterToLinks",
    value: function addTabParameterToLinks() {
      var _this3 = this;
      var currentTab = this.getUrlParameter(this.TAB_PARAM);
      if (!currentTab) {
        var $activeTab = external_jQuery_default()(".nav-tabs .nav-link.active");
        if ($activeTab.length > 0) {
          currentTab = $activeTab.attr("id");
        }
      }
      if (!currentTab) {
        return;
      }

      // Find edit and view links and add the tab parameter
      var linkSelectors = ["#contentview-edit a", "#contentview-view a"];
      external_jQuery_default()(linkSelectors.join(", ")).each(function (index, link) {
        var $link = external_jQuery_default()(link);
        var href = $link.attr("href");
        if (!href || href.indexOf("#") === 0) {
          return;
        }

        // Skip external links
        if (href.indexOf("http") === 0 && href.indexOf(window.location.hostname) === -1) {
          return;
        }

        // Remove existing tab parameter if present
        var regex = new RegExp("([?&])" + _this3.TAB_PARAM + "=.*?(&|$)", "i");
        href = href.replace(regex, "$1").replace(/[?&]$/, "");

        // Add new tab parameter
        var separator = href.indexOf("?") !== -1 ? "&" : "?";
        var newHref = href + separator + _this3.TAB_PARAM + "=" + encodeURIComponent(currentTab);
        $link.attr("href", newHref);
        console.debug("FormTabbing: Updated tab parameter on link: " + newHref);
      });
    }
  }]);
}();
/* harmony default export */ const formtabbing = (FormTabbing);
;// ./senaite.core.js
function senaite_core_createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = senaite_core_unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t["return"] || t["return"](); } finally { if (u) throw o; } } }; }
function senaite_core_unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return senaite_core_arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? senaite_core_arrayLikeToArray(r, a) : void 0; } }
function senaite_core_arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }









document.addEventListener("DOMContentLoaded", function () {
  console.info("*** SENAITE CORE JS LOADED ***");

  // Initialize i18n message factories
  window.i18n = new components_i18n();
  window._t = i18n_wrapper_t;
  window._p = _p;

  // BBB: set global `portal_url` variable
  window.portal_url = document.body.dataset.portalUrl;

  // Initialize Site
  window.site = new site();

  // Initialize SENAITE core namespace
  window.senaite = window.senaite || {};
  window.senaite.core = window.senaite.core || {};

  // Initialize React Sidebar
  window.senaite.core.sidebar = initSidebar();

  // BBB: Keep legacy reference for backwards compatibility
  window.sidebar = window.senaite.core.sidebar;

  // Mount React Workflow Menus (lazy transition lookup on click)
  initWorkflowMenus();

  // Ajax Edit Form Handler
  var form = new editform({
    form_selectors: ["form[name='edit_form']", "form.senaite-ajax-form"],
    field_selectors: ["input[type='text']", "input[type='number']", "input[type='checkbox']", "input[type='radio']", "input[type='file']", "select", "textarea"]
  });
  document.body.addEventListener("datagrid:loaded", function (event) {
    // Init custom CalculationEditForm
    var calculationEditForm = new calculationeditform();
  });

  // Init Tooltips
  external_jQuery_default()(function () {
    external_jQuery_default()("[data-toggle='tooltip']").tooltip();
    external_jQuery_default()("select.selectpicker").selectpicker();
  });

  // Initialize Form Tabbing if tabs are found
  var tabs = document.querySelectorAll(".nav-tabs a[data-toggle='tab']");
  if (tabs.length > 0) {
    var formTabbing = new formtabbing();
    formTabbing.init();
  }

  // Reload the whole view if the status of the view's context has changed
  // due to the transition submission of some items from the listing
  document.body.addEventListener("listing:after_transition_event", function (event) {
    // skip site reload for multi_results view
    var multi_results_templates = ["template-multi_results", "template-multi_results_classic"];
    var body_class_list = document.body.classList;
    for (var _i = 0, _multi_results_templa = multi_results_templates; _i < _multi_results_templa.length; _i++) {
      var class_name = _multi_results_templa[_i];
      if (body_class_list.contains(class_name)) {
        return;
      }
    }

    // get the old workflow state of the view context
    var old_workflow_state = document.body.dataset.reviewState;

    // get the new workflow state of the view context
    // https://github.com/senaite/senaite.app.listing/pull/92
    var config = event.detail.config;
    var new_workflow_state = config.view_context_state;

    // reload the entire page if workflow state of the view context changed
    if (old_workflow_state != new_workflow_state) {
      location.reload();
      return;
    }

    // reload for specific transition and items
    var transition = event.detail.transition;
    var items = event.detail.folderitems;
    var uids = event.detail.uids;
    // filter items that weren't transitioned
    items = items.filter(function (item) {
      return uids.includes(item.uid);
    });
    // iterate over transitioned items and reload if flagged to do so
    var _iterator = senaite_core_createForOfIteratorHelper(items),
      _step;
    try {
      for (_iterator.s(); !(_step = _iterator.n()).done;) {
        var item = _step.value;
        var reload = item.hasOwnProperty("reload") ? item.reload : [];
        if (reload.includes(transition)) {
          // this item was transitioned and flagged with "reload"
          location.reload();
          return;
        }
      }
    } catch (err) {
      _iterator.e(err);
    } finally {
      _iterator.f();
    }
  });

  // BBB: create form Bootstrap navigation tabs for all fieldsets that are
  //      located in a form with the CSS class "enableFormTabbing"
  document.querySelectorAll("form.enableFormTabbing").forEach(function (form) {
    var fieldsets = form.querySelectorAll("fieldset");
    if (fieldsets.length === 0) return;
    var nav = document.createElement("ul");
    nav.className = "nav nav-tabs";
    nav.setAttribute("role", "tablist");
    var tabContent = document.createElement("div");
    tabContent.className = "tab-content";
    fieldsets.forEach(function (fieldset, index) {
      var legend = fieldset.querySelector("legend");
      var tabId = "tab-" + index;
      var li = document.createElement("li");
      li.className = "nav-item";
      var a = document.createElement("a");
      a.className = "nav-link" + (index === 0 ? " active" : "");
      a.setAttribute("data-toggle", "tab");
      a.href = "#" + tabId;
      a.setAttribute("role", "tab");
      a.innerText = legend ? legend.innerText : "Tab " + (index + 1);

      // remove the legend
      legend.remove();
      li.appendChild(a);
      nav.appendChild(li);
      var tabPane = document.createElement("div");
      tabPane.className = "tab-pane fade" + (index === 0 ? " show active" : "");
      tabPane.id = tabId;
      tabPane.setAttribute("role", "tabpanel");
      fieldset.parentNode.insertBefore(tabPane, fieldset);
      tabPane.appendChild(fieldset);
      tabContent.appendChild(tabPane);
    });
    form.insertBefore(nav, form.firstChild);
    form.insertBefore(tabContent, form.firstChild.nextSibling);
  });
});
})();

// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
(() => {
// extracted by mini-css-extract-plugin

})();

/******/ })()
;
//# sourceMappingURL=senaite.core.451275fba835e574c7bf.js.map