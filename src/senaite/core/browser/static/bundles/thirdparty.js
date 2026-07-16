/**
 * @license 
 * jQuery Tools @VERSION Overlay - Overlay base. Extend it.
 * 
 * NO COPYRIGHTS OR LICENSES. DO WHAT YOU LIKE.
 * 
 * http://flowplayer.org/tools/overlay/
 *
 * Since: March 2008
 * Date: @DATE 
 */
(function($) { 

	// static constructs
	$.tools = $.tools || {version: '@VERSION'};
	
	$.tools.overlay = {
		
		addEffect: function(name, loadFn, closeFn) {
			effects[name] = [loadFn, closeFn];	
		},
	
		conf: {  
			close: null,	
			closeOnClick: true,
			closeOnEsc: true,			
			closeSpeed: 'fast',
			effect: 'default',
			
			// since 1.2. fixed positioning not supported by IE6
			fixed: !/msie/.test(navigator.userAgent.toLowerCase()) || navigator.appVersion > 6, 
			
			left: 'center',		
			load: false, // 1.2
			mask: null,  
			oneInstance: true,
			speed: 'normal',
			target: null, // target element to be overlayed. by default taken from [rel]
			top: '10%'
		}
	};

	
	var instances = [], effects = {};
		
	// the default effect. nice and easy!
	$.tools.overlay.addEffect('default', 
		
		/* 
			onLoad/onClose functions must be called otherwise none of the 
			user supplied callback methods won't be called
		*/
		function(pos, onLoad) {
			
			var conf = this.getConf(),
				 w = $(window);				 
				
			if (!conf.fixed)  {
				pos.top += w.scrollTop();
				pos.left += w.scrollLeft();
			} 
				
			pos.position = conf.fixed ? 'fixed' : 'absolute';
			this.getOverlay().css(pos).fadeIn(conf.speed, onLoad); 
			
		}, function(onClose) {
			this.getOverlay().fadeOut(this.getConf().closeSpeed, onClose); 			
		}		
	);		

	
	function Overlay(trigger, conf) {		
		
		// private variables
		var self = this,
			 fire = trigger.add(self),
			 w = $(window), 
			 closers,            
			 overlay,
			 opened,
			 maskConf = $.tools.expose && (conf.mask || conf.expose),
			 uid = Math.random().toString().slice(10);		
		
			 
		// mask configuration
		if (maskConf) {			
			if (typeof maskConf == 'string') { maskConf = {color: maskConf}; }
			maskConf.closeOnClick = maskConf.closeOnEsc = false;
		}			 
		 
		// get overlay and trigger
		var jq = conf.target || trigger.attr("rel");
		overlay = jq ? $(jq) : null || trigger;	
		
		// overlay not found. cannot continue
		if (!overlay.length) { throw "Could not find Overlay: " + jq; }
		
		// trigger's click event
		if (trigger && trigger.index(overlay) == -1) {
			trigger.click(function(e) {				
				self.load(e);
				return e.preventDefault();
			});
		}   			
		
		// API methods  
		$.extend(self, {

			load: function(e) {
				
				// can be opened only once
				if (self.isOpened()) { return self; }
				
				// find the effect
		 		var eff = effects[conf.effect];
		 		if (!eff) { throw "Overlay: cannot find effect : \"" + conf.effect + "\""; }
				
				// close other instances?
				if (conf.oneInstance) {
					$.each(instances, function() {
						this.close(e);
					});
				}
				
				// onBeforeLoad
				e = e || $.Event();
				e.type = "onBeforeLoad";
				fire.trigger(e);				
				if (e.isDefaultPrevented()) { return self; }				

				// opened
				opened = true;
				
				// possible mask effect
				if (maskConf) { $(overlay).expose(maskConf); }				
				
				// position & dimensions 
				var top = conf.top,					
					 left = conf.left,
					 oWidth = overlay.outerWidth(true),
					 oHeight = overlay.outerHeight(true); 
				
				if (typeof top == 'string')  {
					top = top == 'center' ? Math.max((w.height() - oHeight) / 2, 0) : 
						parseInt(top, 10) / 100 * w.height();			
				}				
				
				if (left == 'center') { left = Math.max((w.width() - oWidth) / 2, 0); }

				
		 		// load effect  		 		
				eff[0].call(self, {top: top, left: left}, function() {					
					if (opened) {
						e.type = "onLoad";
						fire.trigger(e);
					}
				}); 				

				// mask.click closes overlay
				if (maskConf && conf.closeOnClick) {
					$.mask.getMask().one("click", self.close); 
				}
				
				// when window is clicked outside overlay, we close
				if (conf.closeOnClick) {
					$(document).on("click." + uid, function(e) { 
						if (!$(e.target).parents(overlay).length) { 
							self.close(e); 
						}
					});						
				}						
			
				// keyboard::escape
				if (conf.closeOnEsc) { 

					// one callback is enough if multiple instances are loaded simultaneously
					$(document).on("keydown." + uid, function(e) {
						if (e.keyCode == 27) { 
							self.close(e);	 
						}
					});			
				}

				
				return self; 
			}, 
			
			close: function(e) {

				if (!self.isOpened()) { return self; }
				
				e = e || $.Event();
				e.type = "onBeforeClose";
				fire.trigger(e);				
				if (e.isDefaultPrevented()) { return; }				
				
				opened = false;
				
				// close effect
				effects[conf.effect][1].call(self, function() {
					e.type = "onClose";
					fire.trigger(e); 
				});
				
				// unbind the keyboard / clicking actions
				$(document).off("click." + uid + " keydown." + uid);		  
				
				if (maskConf) {
					$.mask.close();		
				}
				 
				return self;
			}, 
			
			getOverlay: function() {
				return overlay;	
			},
			
			getTrigger: function() {
				return trigger;	
			},
			
			getClosers: function() {
				return closers;	
			},			

			isOpened: function()  {
				return opened;
			},
			
			// manipulate start, finish and speeds
			getConf: function() {
				return conf;	
			}			
			
		});
		
		// callbacks	
		$.each("onBeforeLoad,onStart,onLoad,onBeforeClose,onClose".split(","), function(i, name) {
				
			// configuration
			if ($.isFunction(conf[name])) { 
				$(self).on(name, conf[name]); 
			}

			// API
			self[name] = function(fn) {
				if (fn) { $(self).on(name, fn); }
				return self;
			};
		});
		
		// close button
		closers = overlay.find(conf.close || ".close");		
		
		if (!closers.length && !conf.close) {
			closers = $('<a class="close"></a>');
			overlay.prepend(closers);	
		}		
		
		closers.click(function(e) { 
			self.close(e);  
		});	
		
		// autoload
		if (conf.load) { self.load(); }
		
	}
	
	// jQuery plugin initialization
	$.fn.overlay = function(conf) {   
		
		// already constructed --> return API
		var el = this.data("overlay");
		if (el) { return el; }	  		 
		
		if ($.isFunction(conf)) {
			conf = {onBeforeLoad: conf};	
		}

		conf = $.extend(true, {}, $.tools.overlay.conf, conf);
		
		this.each(function() {		
			el = new Overlay($(this), conf);
			instances.push(el);
			$(this).data("overlay", el);	
		});
		
		return conf.api ? el: this;		
	}; 
	
})(jQuery);



/**
 * @license 
 * jQuery Tools @VERSION Scrollable - New wave UI design
 * 
 * NO COPYRIGHTS OR LICENSES. DO WHAT YOU LIKE.
 * 
 * http://flowplayer.org/tools/scrollable.html
 *
 * Since: March 2008
 * Date: @DATE 
 */
(function($) { 

	// static constructs
	$.tools = $.tools || {version: '@VERSION'};
	
	$.tools.scrollable = {
		
		conf: {	
			activeClass: 'active',
			circular: false,
			clonedClass: 'cloned',
			disabledClass: 'disabled',
			easing: 'swing',
			initialIndex: 0,
			item: '> *',
			items: '.items',
			keyboard: true,
			mousewheel: false,
			next: '.next',   
			prev: '.prev', 
			size: 1,
			speed: 400,
			vertical: false,
			touch: true,
			wheelSpeed: 0
		} 
	};
					
	// get hidden element's width or height even though it's hidden
	function dim(el, key) {
		var v = parseInt(el.css(key), 10);
		if (v) { return v; }
		var s = el[0].currentStyle; 
		return s && s.width && parseInt(s.width, 10);	
	}

	function find(root, query) { 
		var el = $(query);
		return el.length < 2 ? el : root.parent().find(query);
	}
	
	var current;		
	
	// constructor
	function Scrollable(root, conf) {   
		
		// current instance
		var self = this, 
			 fire = root.add(self),
			 itemWrap = root.children(),
			 index = 0,
			 vertical = conf.vertical;
				
		if (!current) { current = self; } 
		if (itemWrap.length > 1) { itemWrap = $(conf.items, root); }
		
		
		// in this version circular not supported when size > 1
		if (conf.size > 1) { conf.circular = false; } 
		
		// methods
		$.extend(self, {
				
			getConf: function() {
				return conf;	
			},			
			
			getIndex: function() {
				return index;	
			}, 

			getSize: function() {
				return self.getItems().size();	
			},

			getNaviButtons: function() {
				return prev.add(next);	
			},
			
			getRoot: function() {
				return root;	
			},
			
			getItemWrap: function() {
				return itemWrap;	
			},
			
			getItems: function() {
				return itemWrap.find(conf.item).not("." + conf.clonedClass);	
			},
							
			move: function(offset, time) {
				return self.seekTo(index + offset, time);
			},
			
			next: function(time) {
				return self.move(conf.size, time);	
			},
			
			prev: function(time) {
				return self.move(-conf.size, time);	
			},
			
			begin: function(time) {
				return self.seekTo(0, time);	
			},
			
			end: function(time) {
				return self.seekTo(self.getSize() -1, time);	
			},	
			
			focus: function() {
				current = self;
				return self;
			},
			
			addItem: function(item) {
				item = $(item);
				
				if (!conf.circular)  {
					itemWrap.append(item);
					next.removeClass("disabled");
					
				} else {
					itemWrap.children().last().before(item);
					itemWrap.children().first().replaceWith(item.clone().addClass(conf.clonedClass)); 						
				}
				
				fire.trigger("onAddItem", [item]);
				return self;
			},
			
			
			/* all seeking functions depend on this */		
			seekTo: function(i, time, fn) {	
				
				// ensure numeric index
				if (!i.jquery) { i *= 1; }
				
				// avoid seeking from end clone to the beginning
				if (conf.circular && i === 0 && index == -1 && time !== 0) { return self; }
				
				// check that index is sane				
				if (!conf.circular && i < 0 || i > self.getSize() || i < -1) { return self; }
				
				var item = i;
			
				if (i.jquery) {
					i = self.getItems().index(i);	
					
				} else {
					item = self.getItems().eq(i);
				}  
				
				// onBeforeSeek
				var e = $.Event("onBeforeSeek"); 
				if (!fn) {
					fire.trigger(e, [i, time]);				
					if (e.isDefaultPrevented() || !item.length) { return self; }			
				}  
	
				var props = vertical ? {top: -item.position().top} : {left: -item.position().left};  
				
				index = i;
				current = self;  
				if (time === undefined) { time = conf.speed; }   
				
				itemWrap.animate(props, time, conf.easing, fn || function() { 
					fire.trigger("onSeek", [i]);		
				});	 
				
				return self; 
			}					
			
		});
				
		// callbacks	
		$.each(['onBeforeSeek', 'onSeek', 'onAddItem'], function(i, name) {
				
			// configuration
			if ($.isFunction(conf[name])) { 
				$(self).on(name, conf[name]); 
			}
			
			self[name] = function(fn) {
				if (fn) { $(self).on(name, fn); }
				return self;
			};
		});  
		
		// circular loop
		if (conf.circular) {
			
			var cloned1 = self.getItems().slice(-1).clone().prependTo(itemWrap),
				 cloned2 = self.getItems().eq(1).clone().appendTo(itemWrap);

			cloned1.add(cloned2).addClass(conf.clonedClass);
			
			self.onBeforeSeek(function(e, i, time) {
				
				if (e.isDefaultPrevented()) { return; }
				
				/*
					1. animate to the clone without event triggering
					2. seek to correct position with 0 speed
				*/
				if (i == -1) {
					self.seekTo(cloned1, time, function()  {
						self.end(0);		
					});          
					return e.preventDefault();
					
				} else if (i == self.getSize()) {
					self.seekTo(cloned2, time, function()  {
						self.begin(0);		
					});	
				}
				
			});

			// seek over the cloned item

			// if the scrollable is hidden the calculations for seekTo position
			// will be incorrect (eg, if the scrollable is inside an overlay).
			// ensure the elements are shown, calculate the correct position,
			// then re-hide the elements. This must be done synchronously to
			// prevent the hidden elements being shown to the user.

			// See: https://github.com/jquerytools/jquerytools/issues#issue/87

			var hidden_parents = root.parents().add(root).filter(function () {
				if ($(this).css('display') === 'none') {
					return true;
				}
			});
			if (hidden_parents.length) {
				hidden_parents.show();
				self.seekTo(0, 0, function() {});
				hidden_parents.hide();
			}
			else {
				self.seekTo(0, 0, function() {});
			}

		}
		
		// next/prev buttons
		var prev = find(root, conf.prev).click(function(e) { e.stopPropagation(); self.prev(); }),
			 next = find(root, conf.next).click(function(e) { e.stopPropagation(); self.next(); }); 
		
		if (!conf.circular) {
			self.onBeforeSeek(function(e, i) {
				setTimeout(function() {
					if (!e.isDefaultPrevented()) {
						prev.toggleClass(conf.disabledClass, i <= 0);
						next.toggleClass(conf.disabledClass, i >= self.getSize() -1);
					}
				}, 1);
			});
			
			if (!conf.initialIndex) {
				prev.addClass(conf.disabledClass);	
			}			
		}
			
		if (self.getSize() < 2) {
			prev.add(next).addClass(conf.disabledClass);	
		}
			
		// mousewheel support
		if (conf.mousewheel && $.fn.mousewheel) {
			root.mousewheel(function(e, delta)  {
				if (conf.mousewheel) {
					self.move(delta < 0 ? 1 : -1, conf.wheelSpeed || 50);
					return false;
				}
			});			
		}
		
		// touch event
		if (conf.touch) {
			var touch = {};
			
			itemWrap[0].ontouchstart = function(e) {
				var t = e.touches[0];
				touch.x = t.clientX;
				touch.y = t.clientY;
			};
			
			itemWrap[0].ontouchmove = function(e) {
				
				// only deal with one finger
				if (e.touches.length == 1 && !itemWrap.is(":animated")) {			
					var t = e.touches[0],
						 deltaX = touch.x - t.clientX,
						 deltaY = touch.y - t.clientY;
	
					self[vertical && deltaY > 0 || !vertical && deltaX > 0 ? 'next' : 'prev']();				
					e.preventDefault();
				}
			};
		}
		
		if (conf.keyboard)  {
			
			$(document).on("keydown.scrollable", function(evt) {

				// skip certain conditions
				if (!conf.keyboard || evt.altKey || evt.ctrlKey || evt.metaKey || $(evt.target).is(":input")) { 
					return; 
				}
				
				// does this instance have focus?
				if (conf.keyboard != 'static' && current != self) { return; }
					
				var key = evt.keyCode;
			
				if (vertical && (key == 38 || key == 40)) {
					self.move(key == 38 ? -1 : 1);
					return evt.preventDefault();
				}
				
				if (!vertical && (key == 37 || key == 39)) {					
					self.move(key == 37 ? -1 : 1);
					return evt.preventDefault();
				}	  
				
			});  
		}
		
		// initial index
		if (conf.initialIndex) {
			self.seekTo(conf.initialIndex, 0, function() {});
		}
	} 

		
	// jQuery plugin implementation
	$.fn.scrollable = function(conf) { 
			
		// already constructed --> return API
		var el = this.data("scrollable");
		if (el) { return el; }		 

		conf = $.extend({}, $.tools.scrollable.conf, conf); 
		
		this.each(function() {			
			el = new Scrollable($(this), conf);
			$(this).data("scrollable", el);	
		});
		
		return conf.api ? el: this; 
		
	};
			
	
})(jQuery);


/**
 * @license 
 * jQuery Tools @VERSION Tabs- The basics of UI design.
 * 
 * NO COPYRIGHTS OR LICENSES. DO WHAT YOU LIKE.
 * 
 * http://flowplayer.org/tools/tabs/
 *
 * Since: November 2008
 * Date: @DATE 
 */  
(function($) {
		
	// static constructs
	$.tools = $.tools || {version: '@VERSION'};
	
	$.tools.tabs = {
		
		conf: {
			tabs: 'a',
			current: 'current',
			onBeforeClick: null,
			onClick: null, 
			effect: 'default',
			initialEffect: false,   // whether or not to show effect in first init of tabs
			initialIndex: 0,			
			event: 'click',
			rotate: false,
			
      // slide effect
      slideUpSpeed: 400,
      slideDownSpeed: 400,
			
			// 1.2
			history: false
		},
		
		addEffect: function(name, fn) {
			effects[name] = fn;
		}
		
	};
	
	var effects = {
		
		// simple "toggle" effect
		'default': function(i, done) { 
			this.getPanes().hide().eq(i).show();
			done.call();
		}, 
		
		/*
			configuration:
				- fadeOutSpeed (positive value does "crossfading")
				- fadeInSpeed
		*/
		fade: function(i, done) {		
			
			var conf = this.getConf(),
				 speed = conf.fadeOutSpeed,
				 panes = this.getPanes();
			
			if (speed) {
				panes.fadeOut(speed);	
			} else {
				panes.hide();	
			}

			panes.eq(i).fadeIn(conf.fadeInSpeed, done);	
		},
		
		// for basic accordions
		slide: function(i, done) {
		  var conf = this.getConf();
		  
			this.getPanes().slideUp(conf.slideUpSpeed);
			this.getPanes().eq(i).slideDown(conf.slideDownSpeed, done);			 
		}, 

		/**
		 * AJAX effect
		 */
		ajax: function(i, done)  {			
			this.getPanes().eq(0).load(this.getTabs().eq(i).attr("href"), done);	
		}		
	};   	
	
	/**
	 * Horizontal accordion
	 * 
	 * @deprecated will be replaced with a more robust implementation
	*/
	
	var
	  /**
	  *   @type {Boolean}
	  *
	  *   Mutex to control horizontal animation
	  *   Disables clicking of tabs while animating
	  *   They mess up otherwise as currentPane gets set *after* animation is done
	  */
	  animating,
	  /**
	  *   @type {Number}
	  *   
	  *   Initial width of tab panes
	  */
	  w;
	 
	$.tools.tabs.addEffect("horizontal", function(i, done) {
	  if (animating) return;    // don't allow other animations
	  
	  var nextPane = this.getPanes().eq(i),
	      currentPane = this.getCurrentPane();
	      
		// store original width of a pane into memory
		w || ( w = this.getPanes().eq(0).width() );
		animating = true;
		
		nextPane.show(); // hidden by default
		
		// animate current pane's width to zero
    // animate next pane's width at the same time for smooth animation
    currentPane.animate({width: 0}, {
      step: function(now){
        nextPane.css("width", w-now);
      },
      complete: function(){
        $(this).hide();
        done.call();
        animating = false;
     }
    });
    // Dirty hack...  onLoad, currentPant will be empty and nextPane will be the first pane
    // If this is the case, manually run callback since the animation never occured, and reset animating
    if (!currentPane.length){ 
      done.call(); 
      animating = false;
    }
	});	

	
	function Tabs(root, paneSelector, conf) {
		
		var self = this,
        trigger = root.add(this),
        tabs = root.find(conf.tabs),
        panes = paneSelector.jquery ? paneSelector : root.children(paneSelector),
        current;
			 
		
		// make sure tabs and panes are found
		if (!tabs.length)  { tabs = root.children(); }
		if (!panes.length) { panes = root.parent().find(paneSelector); }
		if (!panes.length) { panes = $(paneSelector); }
		
		
		// public methods
		$.extend(this, {				
			click: function(i, e) {
			  
				var tab = tabs.eq(i),
				    firstRender = !root.data('tabs');
				
				if (typeof i == 'string' && i.replace("#", "")) {
					tab = tabs.filter("[href*=\"" + i.replace("#", "") + "\"]");
					i = Math.max(tabs.index(tab), 0);
				}
								
				if (conf.rotate) {
					var last = tabs.length -1; 
					if (i < 0) { return self.click(last, e); }
					if (i > last) { return self.click(0, e); }						
				}
				
				if (!tab.length) {
					if (current >= 0) { return self; }
					i = conf.initialIndex;
					tab = tabs.eq(i);
				}				
				
				// current tab is being clicked
				if (i === current) { return self; }
				
				// possibility to cancel click action				
				e = e || $.Event();
				e.type = "onBeforeClick";
				trigger.trigger(e, [i]);				
				if (e.isDefaultPrevented()) { return; }
				
        // if firstRender, only run effect if initialEffect is set, otherwise default
				var effect = firstRender ? conf.initialEffect && conf.effect || 'default' : conf.effect;

				// call the effect
				effects[effect].call(self, i, function() {
					current = i;
					// onClick callback
					e.type = "onClick";
					trigger.trigger(e, [i]);
				});			
				
				// default behaviour
				tabs.removeClass(conf.current);	
				tab.addClass(conf.current);				
				
				return self;
			},
			
			getConf: function() {
				return conf;	
			},

			getTabs: function() {
				return tabs;	
			},
			
			getPanes: function() {
				return panes;	
			},
			
			getCurrentPane: function() {
				return panes.eq(current);	
			},
			
			getCurrentTab: function() {
				return tabs.eq(current);	
			},
			
			getIndex: function() {
				return current;	
			}, 
			
			next: function() {
				return self.click(current + 1);
			},
			
			prev: function() {
				return self.click(current - 1);	
			},
			
			destroy: function() {
				tabs.off(conf.event).removeClass(conf.current);
				panes.find("a[href^=\"#\"]").off("click.T"); 
				return self;
			}
		
		});

		// callbacks	
		$.each("onBeforeClick,onClick".split(","), function(i, name) {
				
			// configuration
			if ($.isFunction(conf[name])) {
				$(self).on(name, conf[name]); 
			}

			// API
			self[name] = function(fn) {
				if (fn) { $(self).on(name, fn); }
				return self;	
			};
		});
	
		
		if (conf.history && $.fn.history) {
			$.tools.history.init(tabs);
			conf.event = 'history';
		}	
		
		// setup click actions for each tab
		tabs.each(function(i) { 				
			$(this).on(conf.event, function(e) {
				self.click(i, e);
				return e.preventDefault();
			});			
		});
		
		// cross tab anchor link
		panes.find("a[href^=\"#\"]").on("click.T", function(e) {
			self.click($(this).attr("href"), e);		
		}); 
		
		// open initial tab
		if (location.hash && conf.tabs == "a" && root.find("[href=\"" +location.hash+ "\"]").length) {
			self.click(location.hash);

		} else {
			if (conf.initialIndex === 0 || conf.initialIndex > 0) {
				self.click(conf.initialIndex);
			}
		}				
		
	}
	
	
	// jQuery plugin implementation
	$.fn.tabs = function(paneSelector, conf) {
		
		// return existing instance
		var el = this.data("tabs");
		if (el) { 
			el.destroy();	
			this.removeData("tabs");
		}

		if ($.isFunction(conf)) {
			conf = {onBeforeClick: conf};
		}
		
		// setup conf
		conf = $.extend({}, $.tools.tabs.conf, conf);		
		
		
		this.each(function() {				
			el = new Tabs($(this), paneSelector, conf);
			$(this).data("tabs", el); 
		});		
		
		return conf.api ? el: this;		
	};		
		
}) (jQuery); 




/**
 * @license 
 * jQuery Tools @VERSION History "Back button for AJAX apps"
 * 
 * NO COPYRIGHTS OR LICENSES. DO WHAT YOU LIKE.
 * 
 * http://flowplayer.org/tools/toolbox/history.html
 * 
 * Since: Mar 2010
 * Date: @DATE 
 */
(function($) {
		
	var hash, iframe, links, inited;		
	
	$.tools = $.tools || {version: '@VERSION'};
	
	$.tools.history = {
	
		init: function(els) {
			
			if (inited) { return; }
			
			// IE
			if ($.browser.msie && $.browser.version < '8') {
				
				// create iframe that is constantly checked for hash changes
				if (!iframe) {
					iframe = $("<iframe/>").attr("src", "javascript:false;").hide().get(0);
					$("body").append(iframe);
									
					setInterval(function() {
						var idoc = iframe.contentWindow.document, 
							 h = idoc.location.hash;
					
						if (hash !== h) {						
							$(window).trigger("hash", h);
						}
					}, 100);
					
					setIframeLocation(location.hash || '#');
				}

				
			// other browsers scans for location.hash changes directly without iframe hack
			} else { 
				setInterval(function() {
					var h = location.hash;
					if (h !== hash) {
						$(window).trigger("hash", h);
					}						
				}, 100);
			}

			links = !links ? els : links.add(els);
			
			els.click(function(e) {
				var href = $(this).attr("href");
				if (iframe) { setIframeLocation(href); }
				
				// handle non-anchor links
				if (href.slice(0, 1) != "#") {
					location.href = "#" + href;
					return e.preventDefault();		
				}
				
			}); 
			
			inited = true;
		}	
	};  
	

	function setIframeLocation(h) {
		if (h) {
			var doc = iframe.contentWindow.document;
			doc.open().close();	
			doc.location.hash = h;
		}
	} 
		 
	// global histroy change listener
	$(window).on("hash", function(e, h)  { 
		if (h) {
			links.filter(function() {
			  var href = $(this).attr("href");
			  return href == h || href == h.replace("#", ""); 
			}).trigger("history", [h]);	
		} else {
			links.eq(0).trigger("history", [h]);	
		}

		hash = h;

	});
		
	
	// jQuery plugin implementation
	$.fn.history = function(fn) {
			
		$.tools.history.init(this);

		// return jQuery
		return this.on("history", fn);		
	};	
		
})(jQuery); 



/**
 * @license 
 * jQuery Tools @VERSION / Expose - Dim the lights
 * 
 * NO COPYRIGHTS OR LICENSES. DO WHAT YOU LIKE.
 * 
 * http://flowplayer.org/tools/toolbox/expose.html
 *
 * Since: Mar 2010
 * Date: @DATE 
 */
(function($) { 	

	// static constructs
	$.tools = $.tools || {version: '@VERSION'};
	
	var tool;
	
	tool = $.tools.expose = {
		
		conf: {	
			maskId: 'exposeMask',
			loadSpeed: 'slow',
			closeSpeed: 'fast',
			closeOnClick: true,
			closeOnEsc: true,
			
			// css settings
			zIndex: 9998,
			opacity: 0.8,
			startOpacity: 0,
			color: '#fff',
			
			// callbacks
			onLoad: null,
			onClose: null
		}
	};

	/* one of the greatest headaches in the tool. finally made it */
	function viewport() {
				
		// the horror case
		if (/msie/.test(navigator.userAgent.toLowerCase())) {
			
			// if there are no scrollbars then use window.height
			var d = $(document).height(), w = $(window).height();
			
			return [
				window.innerWidth || 							// ie7+
				document.documentElement.clientWidth || 	// ie6  
				document.body.clientWidth, 					// ie6 quirks mode
				d - w < 20 ? w : d
			];
		} 
		
		// other well behaving browsers
		return [$(document).width(), $(document).height()]; 
	} 
	
	function call(fn) {
		if (fn) { return fn.call($.mask); }
	}
	
	var mask, exposed, loaded, config, overlayIndex;		
	
	
	$.mask = {
		
		load: function(conf, els) {
			
			// already loaded ?
			if (loaded) { return this; }			
			
			// configuration
			if (typeof conf == 'string') {
				conf = {color: conf};	
			}
			
			// use latest config
			conf = conf || config;
			
			config = conf = $.extend($.extend({}, tool.conf), conf);

			// get the mask
			mask = $("#" + conf.maskId);
				
			// or create it
			if (!mask.length) {
				mask = $('<div/>').attr("id", conf.maskId);
				$("body").append(mask);
			}
			
			// set position and dimensions 			
			var size = viewport();
				
			mask.css({				
				position:'absolute', 
				top: 0, 
				left: 0,
				width: size[0],
				height: size[1],
				display: 'none',
				opacity: conf.startOpacity,					 		
				zIndex: conf.zIndex 
			});
			
			if (conf.color) {
				mask.css("backgroundColor", conf.color);	
			}			
			
			// onBeforeLoad
			if (call(conf.onBeforeLoad) === false) {
				return this;
			}
			
			// esc button
			if (conf.closeOnEsc) {						
				$(document).on("keydown.mask", function(e) {							
					if (e.keyCode == 27) {
						$.mask.close(e);	
					}		
				});			
			}
			
			// mask click closes
			if (conf.closeOnClick) {
				mask.on("click.mask", function(e)  {
					$.mask.close(e);		
				});					
			}			
			
			// resize mask when window is resized
			$(window).on("resize.mask", function() {
				$.mask.fit();
			});
			
			// exposed elements
			if (els && els.length) {
				
				overlayIndex = els.eq(0).css("zIndex");

				// make sure element is positioned absolutely or relatively
				$.each(els, function() {
					var el = $(this);
					if (!/relative|absolute|fixed/i.test(el.css("position"))) {
						el.css("position", "relative");		
					}					
				});
			 
				// make elements sit on top of the mask
				exposed = els.css({ zIndex: Math.max(conf.zIndex + 1, overlayIndex == 'auto' ? 0 : overlayIndex)});			
			}	
			
			// reveal mask
			mask.css({display: 'block'}).fadeTo(conf.loadSpeed, conf.opacity, function() {
				$.mask.fit(); 
				call(conf.onLoad);
				loaded = "full";
			});
			
			loaded = true;			
			return this;				
		},
		
		close: function() {
			if (loaded) {
				
				// onBeforeClose
				if (call(config.onBeforeClose) === false) { return this; }
					
				mask.fadeOut(config.closeSpeed, function()  {										
					if (exposed) {
						exposed.css({zIndex: overlayIndex});						
					}				
					loaded = false;
					call(config.onClose);
				});				
				
				// unbind various event listeners
				$(document).off("keydown.mask");
				mask.off("click.mask");
				$(window).off("resize.mask");  
			}
			
			return this; 
		},
		
		fit: function() {
			if (loaded) {
				var size = viewport();				
				mask.css({width: size[0], height: size[1]});
			}				
		},
		
		getMask: function() {
			return mask;	
		},
		
		isLoaded: function(fully) {
			return fully ? loaded == 'full' : loaded;	
		}, 
		
		getConf: function() {
			return config;	
		},
		
		getExposed: function() {
			return exposed;	
		}		
	};
	
	$.fn.mask = function(conf) {
		$.mask.load(conf);
		return this;		
	};			
	
	$.fn.expose = function(conf) {
		$.mask.load(conf, this);
		return this;			
	};


})(jQuery);

/*****************

   SENAITE
   -------

   - Removed `buildFragment` function call
     https://github.com/senaite/senaite.lims/issues/21


   jQuery Tools overlay helpers.

   Copyright © 2010, The Plone Foundation
   Licensed under the GPL, see LICENSE.txt for details.

*****************/

/*jslint browser: true */
/*global jQuery, ajax_noresponse_message, close_box_message, window */

// Name space object for pipbox
var pb = {spinner: {}, overlay_counter: 1};

jQuery.tools.overlay.conf.oneInstance = false;

(function($) {
    // override jqt default effect to take account of position
    // of parent elements
    jQuery.tools.overlay.addEffect('default',
        function(pos, onLoad) {
            var conf = this.getConf(),
                 w = $(window),
                 ovl = this.getOverlay(),
                 op = ovl.parent().offsetParent().offset();

            if (!conf.fixed)  {
                pos.top += w.scrollTop() - op.top;
                pos.left += w.scrollLeft() - op.left;
            }

            pos.position = conf.fixed ? 'fixed' : 'absolute';
            ovl.css(pos).fadeIn(conf.speed, onLoad);

        }, function(onClose) {
            this.getOverlay().fadeOut(this.getConf().closeSpeed, onClose);
        }
    );

    pb.spinner.show = function () {
        $('body').css('cursor', 'wait');
    };
    pb.spinner.hide = function () {
        $('body').css('cursor', '');
    };
}(jQuery));


jQuery(function ($) {

    /******
        $.fn.prepOverlay jQuery plugin to inject overlay target into DOM and
        annotate it with the data we'll need in order to display it.
    ******/
    $.fn.prepOverlay = function (pba) {
        return this.each(function () {
            var o, pbo, config, onBeforeLoad, onLoad, src, parts;

            o = $(this);

            // copy options so that it's not just a reference
            // to the parameter.
            pbo = $.extend(true, {'width':'60%'}, pba);

            // set overlay configuration
            config = pbo.config || {};

            // set onBeforeLoad handler
            onBeforeLoad = pb[pbo.subtype];
            if (onBeforeLoad) {
                config.onBeforeLoad = onBeforeLoad;
            }
            onLoad = config.onLoad;
            config.onLoad = function () {
                if (onLoad) {
                    onLoad.apply(this, arguments);
                }
                pb.fi_focus(this.getOverlay());
            };

            // be promiscuous, pick up the url from
            // href, src or action attributes
            src = o.attr('href') || o.attr('src') || o.attr('action');

            // translate url with config specifications
            if (pbo.urlmatch) {
                src = src.replace(new RegExp(pbo.urlmatch), pbo.urlreplace);
            }

            if (pbo.subtype === 'inline') {
                // we're going to let tools' overlay do all the real
                // work. Just get the markers in place.
                src = src.replace(/^.+#/, '#');
                $("[id='" + src.replace('#', '') + "']").addClass('overlay');
                o.removeAttr('href').attr('rel', src);
                // use overlay on the source (clickable) element
                o.overlay();
            } else {
                // save various bits of information from the pbo options,
                // and enable the overlay.

                // this is not inline, so in one fashion or another
                // we'll be loading it via the beforeLoad callback.
                // create a unique id for a target element
                pbo.nt = 'pb_' + pb.overlay_counter;
                pb.overlay_counter += 1;

                pbo.selector = pbo.filter || pbo.selector;
                if (!pbo.selector) {
                    // see if one's been supplied in the src
                    parts = src.split(' ');
                    src = parts.shift();
                    pbo.selector = parts.join(' ');
                }

                pbo.src = src;
                pbo.config = config;
                pbo.source = o;

                // remove any existing overlay and overlay handler
                pb.remove_overlay(o);

                // save options on trigger element
                o.data('pbo', pbo);

                // mark the source with a rel attribute so we can find
                // the overlay, and a special class for styling
                o.attr('rel', '#' + pbo.nt);
                o.addClass('link-overlay');

                // for some subtypes, we're setting click handlers
                // and attaching overlay to the target element. That's
                // so we'll know the dimensions early.
                // Others, like iframe, just use overlay.

                register_handler = function(evt, o, hnd) {
                   switch (evt) {
                      case undefined:
                         o.click(hnd);
                         break;
                      case 'click':
                         o.click(hnd);
                         break;
                      case 'hover':
                         o.hover(hnd);
                         break;
                      case 'dblclick':
                         o.dblclick(hnd);
                         break;
                      default:
                         throw "Unsupported event";
                   }
                }

                switch (pbo.subtype) {
                case 'image':
                    register_handler(pbo.event, o, pb.image_click);
                    break;
                case 'ajax':
                    register_handler(pbo.event, o, pb.ajax_click);
                    break;
                case 'iframe':
                    pb.create_content_div(pbo);
                    o.overlay(config);
                    break;
                default:
                    throw "Unsupported overlay type";
                }

                // in case the click source wasn't
                // already a link.
                o.css('cursor', 'pointer');
            }
        });
    };


    /******
        pb.remove_overlay
        Remove the overlay and handler associated with a jquery wrapped
        trigger object
    ******/
    pb.remove_overlay = function (o) {
        var old_data = o.data('pbo');
        if (old_data) {
            switch (old_data.subtype) {
                case 'image':
                    o.unbind('click', pb.image_click);
                    break;
                case 'ajax':
                    o.unbind('click', pb.ajax_click);
                    break;
                default:
                    // it's probably the jqt overlay click handler,
                    // but we don't know the handler and are forced
                    // to do a generic unbind of click handlers.
                    o.unbind('click');
                    break;
            }
            if (old_data.nt) {
                $('#' + old_data.nt).remove();
            }
        }
    };


    /******
        pb.create_content_div
        create a div to act as an overlay; append it to
        the body; return it
    ******/
    pb.create_content_div = function (pbo, trigger) {
        var content,
            close_message,
            pbw = pbo.width;

        if (typeof close_box_message === 'undefined') {
            close_message = 'Close this box.';
        } else {
            close_message = close_box_message;
        }

        content = $(
            '<div id="' + pbo.nt +
            '" class="overlay overlay-' + pbo.subtype +
            ' ' + (pbo.cssclass || '') +
            '"><div class="close"><a href="#" class="hiddenStructure" title="Close this box">' +
            close_message +
            '</a></div></div>');

        content.data('pbo', pbo);

        // if a width option is specified, set it on the overlay div,
        // computing against the window width if a % was specified.
        if (pbw) {
            if (pbw.indexOf('%') > 0) {
                content.width(parseInt(pbw, 10) / 100 * $(window).width());
            } else {
                content.width(pbw);
            }
        }

        // add the target element at the end of the body.
        content.appendTo($("body"));

        return content;
    };


    /******
        pb.image_click
        click handler for image loads
    ******/
    pb.image_click = function (event) {
        var ethis, content, api, img, el, pbo;

        ethis = $(this);
        pbo = ethis.data('pbo');

        // find target container
        content = $(ethis.attr('rel'));
        if (!content.length) {
            content = pb.create_content_div(pbo);
            content.overlay(pbo.config);
        }
        api = content.overlay();

        // is the image loaded yet?
        if (content.find('img').length === 0) {
            // load the image.
            if (pbo.src) {
                pb.spinner.show();

                // create the image and stuff it
                // into our target
                img = new Image();
                img.src = pbo.src;
                el = $(img);
                content.append(el.addClass('pb-image'));

                // Now, we'll cause the overlay to
                // load when the image is loaded.
                el.load(function () {
                    pb.spinner.hide();
                    api.load();
                });

            }
        } else {
            api.load();
        }

        return false;
    };


    /******
        pb.fi_focus
        First-input focus inside $ selection.
    ******/
    pb.fi_focus = function (jqo) {
        if (! jqo.find("form div.error :input:first").focus().length) {
            jqo.find("form :input:visible:first").focus();
        }
    };


    /******
        pb.ajax_error_recover
        jQuery's ajax load function does not load error responses.
        This routine returns the cooked error response.
    ******/
    pb.ajax_error_recover = function (responseText, selector) {
        var tcontent = $('<div/>').append(
                responseText.replace(/<script(.|\s)*?\/script>/gi, ""));
        return selector ? tcontent.find(selector) : tcontent;
    };


    /******
        pb.add_ajax_load
        Adds a hidden ajax_load input to form
    ******/
    pb.add_ajax_load = function (form) {
        if (form.find('input[name=ajax_load]').length === 0) {
            form.prepend($('<input type="hidden" name="ajax_load" value="' +
                (new Date().getTime()) +
                '" />'));
        }
    };

    /******
        pb.prep_ajax_form
        Set up form with ajaxForm, including success and error handlers.
    ******/
    pb.prep_ajax_form = function (form) {
        var ajax_parent = form.closest('.pb-ajax'),
            data_parent = ajax_parent.closest('.overlay-ajax'),
            pbo = data_parent.data('pbo'),
            formtarget = pbo.formselector,
            closeselector = pbo.closeselector,
            beforepost = pbo.beforepost,
            afterpost = pbo.afterpost,
            noform = pbo.noform,
            api = data_parent.overlay(),
            selector = pbo.selector,
            options = {};

        options.beforeSerialize = function () {
            pb.spinner.show();
        };

        if (beforepost) {
            options.beforeSubmit = function (arr, form, options) {
                return beforepost(form, arr, options);
            };
        }
        options.success = function (responseText, statusText, xhr, form) {
            $(document).trigger('formOverlayStart', [this, responseText, statusText, xhr, form]);
            // success comes in many forms, some of which are errors;
            //

            var el, myform, success, target, scripts = [], filteredResponseText;

            success = statusText === "success" || statusText === "notmodified";

            if (! success) {
                // The responseText parameter is actually xhr
                responseText = responseText.responseText;
            }
            // strip inline script tags
            filteredResponseText = responseText.replace(/<script(.|\s)*?\/script>/gi, "");

            // create a div containing the optionally filtered response
            el = $('<div />').append(
                selector ?
                    // a lesson learned from the jQuery source: $(responseText)
                    // will not work well unless responseText is well-formed;
                    // appending to a div is more robust, and automagically
                    // removes the html/head/body outer tagging.
                    $('<div />').append(filteredResponseText).find(selector)
                    :
                    filteredResponseText
                );

            // afterpost callback
            if (success && afterpost) {
                afterpost(el, data_parent);
            }

            myform = el.find(formtarget);
            if (success && myform.length) {
                ajax_parent.empty().append(el);

                // execute inline scripts

                // try {
                //     // jQuery 1.7
                //     $.buildFragment([responseText], [document], scripts);
                // } catch (e) {
                //     // jQuery 1.9
                //     $.buildFragment([responseText], document, scripts);
                // }

                if (scripts.length) {
                    $.each(scripts, function() {
                        $.globalEval( this.text || this.textContent || this.innerHTML || "" );
                    });
                }

                // This may be a complex form.
                if ($.fn.ploneTabInit) {
                    el.ploneTabInit();
                }
                pb.fi_focus(ajax_parent);

                pb.add_ajax_load(myform);
                // attach submit handler with the same options
                myform.ajaxForm(options);

                // attach close to element id'd by closeselector
                if (closeselector) {
                    el.find(closeselector).click(function (event) {
                        api.close();
                        return false;
                    });
                }
                $(document).trigger('formOverlayLoadSuccess', [this, myform, api, pb, ajax_parent]);
            } else {
                // there's no form in our new content or there's been an error
                if (success) {
                    if (typeof noform === "function") {
                        // get action from callback
                        noform = noform(el, pbo);
                    }
                } else {
                    noform = statusText;
                }


                switch (noform) {
                case 'close':
                    api.close();
                    break;
                case 'reload':
                    api.close();
                    pb.spinner.show();
                    // location.reload results in a repost
                    // dialog in some browsers; very unlikely to
                    // be what we want.
                    location.replace(location.href);
                    break;
                case 'redirect':
                    api.close();
                    pb.spinner.show();
                    target = pbo.redirect;
                    if (typeof target === "function") {
                        // get target from callback
                        target = target(el, responseText, pbo);
                    }
                    location.replace(target);
                    break;
                default:
                    if (el.children()) {
                        // show what we've got
                        ajax_parent.empty().append(el);
                    } else {
                        api.close();
                    }
                    break;
                }
                $(document).trigger('formOverlayLoadFailure', [this, myform, api, pb, ajax_parent, noform]);
            }
            pb.spinner.hide();
        };
        // error and success callbacks are the same
        options.error = options.success;

        pb.add_ajax_load(form);
        form.ajaxForm(options);
    };


    /******
        pb.ajax_click
        Click handler for ajax sources. The job of this routine
        is to do the ajax load of the overlay element, then
        call the JQT overlay loader.
    ******/
    pb.ajax_click = function (event) {
        var ethis = $(this),
            pbo,
            content,
            api,
            src,
            el,
            selector,
            formtarget,
            closeselector,
            sep,
            scripts = [],
            e;

        e = $.Event();
        e.type = "beforeAjaxClickHandled";
        $(document).trigger(e, [this, event]);
        if (e.isDefaultPrevented()) { return false; }

        pbo = ethis.data('pbo');

        content = pb.create_content_div(pbo, ethis);
        // pbo.config.top = $(window).height() * 0.1 - ethis.offsetParent().offset().top;
        content.overlay(pbo.config, ethis);
        api = content.overlay();
        src = pbo.src;
        selector = pbo.selector;
        formtarget = pbo.formselector;
        closeselector = pbo.closeselector;

        pb.spinner.show();

        // prevent double click warning for this form
        $(this).find("input.submitting").removeClass('submitting');

        el = $('div.pb-ajax', content);
        if (el.length === 0) {
            el = $('<div class="pb-ajax" />');
            content.append(el);
        }
        if (api.getConf().fixed) {
            // don't let it be over 75% of the viewport's height
            el.css('max-height', Math.floor($(window).height() * 0.75));
        }

        // affix a random query argument to prevent
        // loading from browser cache
        sep = (src.indexOf('?') >= 0) ? '&': '?';
        src += sep + "ajax_load=" + (new Date().getTime());

        // add selector, if any
        if (selector) {
            src += ' ' + selector;
        }

        // set up callback to be used whenever new contents are loaded
        // into the overlay, to prepare links and forms to stay within
        // the overlay
        el[0].handle_load_inside_overlay = function(responseText, errorText) {
            var ele, target;
            ele = $(this);

            if (errorText === 'error') {
                ele.append(pb.ajax_error_recover(responseText, selector));
            } else if (!responseText.length) {
                ele.append(ajax_noresponse_message || 'No response from server.');
            }

            // a non-semantic div here will make sure we can
            // do enough formatting.
            ele.wrapInner('<div />');

            // add the submit handler if we've a formtarget
            if (formtarget) {
                target = ele.find(formtarget);
                if (target.length > 0) {
                    pb.prep_ajax_form(target);
                }
            }

            // if a closeselector has been specified, tie it to the overlay's
            // close method via closure
            if (closeselector) {
                ele.find(closeselector).click(function (event) {
                    api.close();
                    return false;
                });
            }

            // execute inline scripts
            // try {
            //     // jQuery 1.7
            //     $.buildFragment([responseText], [document], scripts);
            // } catch (e) {
            //     // jQuery 1.9
            //     debugger
            //     $.buildFragment([responseText], document, scripts);
            // }

            if (scripts.length) {
                $.each(scripts, function() {
                    $.globalEval( this.text || this.textContent || this.innerHTML || "" );
                });
            }

            // This may be a complex form.
            if ($.fn.ploneTabInit) {
                ele.ploneTabInit();
            }

            // remove element on close so that it doesn't congest the DOM
            api.onClose = function () {
                content.remove();
            };
            $(document).trigger('loadInsideOverlay', [this, responseText, errorText, api]);
        };

        // and load the div
        var data = null;
        var form = $(pbo.source.context);
        var method = form.attr('method');
        if (method && method.toLowerCase() == 'post') {
            data = form.serializeArray();
        }
        // load will do a POST instead of GET if data is not null
        el.load(src, data, function (responseText, errorText) {
            // post-process the overlay contents
            el[0].handle_load_inside_overlay.apply(this, [responseText, errorText]);

            // Now, it's all ready to display; hide the
            // spinner and call JQT overlay load.
            pb.spinner.hide();
            api.load();

            return true;
        });

        // don't do the default action
        return false;
    };


    /******
        pb.iframe
        onBeforeLoad handler for iframe overlays.

        Note that the spinner is handled a little differently
        so that we can keep it displayed while the iframe's
        content is loading.
    ******/
    pb.iframe = function () {
        var content, pbo;

        pb.spinner.show();

        content = this.getOverlay();
        pbo = this.getTrigger().data('pbo');

        if (content.find('iframe').length === 0 && pbo.src) {
            content.append(
                '<iframe src="' + pbo.src + '" width="' +
                 content.width() + '" height="' + content.height() +
                 '" onload="pb.spinner.hide()"/>');
        } else {
            pb.spinner.hide();
        }
        return true;
    };

    // $('.newsImageContainer a')
    //     .prepOverlay({
    //          subtype: 'image',
    //          urlmatch: '/image_view_fullscreen$',
    //          urlreplace: '_preview'
    //         });

});

/*
 *  BarCode Coder Library (BCC Library)
 *  BCCL Version 2.0
 *
 *  Porting : jQuery barcode plugin
 *  Version : 2.2
 *
 *  Date  : 2019-02-21
 *  Author  : DEMONTE Jean-Baptiste <jbdemonte@gmail.com>
 *            HOUREZ Jonathan
 *
 *  Web site: http://barcode-coder.com/
 *  dual licence :  http://www.cecill.info/licences/Licence_CeCILL_V2-fr.html
 *                  http://www.gnu.org/licenses/gpl.html
 */
!function(g){var s={barWidth:1,barHeight:50,moduleSize:5,showHRI:!0,addQuietZone:!0,marginHRI:5,bgColor:"#FFFFFF",color:"#000000",fontSize:10,output:"css",posX:0,posY:0},h=["NNWWN","WNNNW","NWNNW","WWNNN","NNWNW","WNWNN","NWWNN","NNNWW","WNNWN","NWNWN"];function d(r,t,e){var n,o,a,i,c,f="";if(r=function(r,t,e){var n,o,a=!0,i=0;if(t){for("int25"===e&&r.length%2==0&&(r="0"+r),n=r.length-1;-1<n;n--){if(o=R(r.charAt(n)),isNaN(o))return"";i+=a?3*o:o,a=!a}r+=((10-i%10)%10).toString()}else r.length%2!=0&&(r="0"+r);return r}(r,t,e),"int25"===e){for(f+="1010",n=0;n<r.length/2;n++)for(a=r.charAt(2*n),i=r.charAt(2*n+1),o=0;o<5;o++)f+="1","W"===h[a].charAt(o)&&(f+="1"),f+="0","W"===h[i].charAt(o)&&(f+="0");f+="1101"}else if("std25"===e){for(f+="11011010",n=0;n<r.length;n++)for(c=r.charAt(n),o=0;o<5;o++)f+="1","W"===h[c].charAt(o)&&(f+="11"),f+="0";f+="11010110"}return f}var v=[["0001101","0100111","1110010"],["0011001","0110011","1100110"],["0010011","0011011","1101100"],["0111101","0100001","1000010"],["0100011","0011101","1011100"],["0110001","0111001","1001110"],["0101111","0000101","1010000"],["0111011","0010001","1000100"],["0110111","0001001","1001000"],["0001011","0010111","1110100"]],A=["000000","001011","001101","001110","010011","011001","011100","010101","010110","011010"],p=["00","01","10","11"],b=["11000","10100","10010","10001","01100","00110","00011","01010","01001","00101"];function m(r,t){var e,n,o,a,i,c,f,h,l,u,g="ean8"===t?7:12,s=r;if(!(r=r.substring(0,g)).match(new RegExp("^[0-9]{"+g+"}$")))return"";if(r=N(r,t),c="101","ean8"===t){for(e=0;e<4;e++)c+=v[R(r.charAt(e))][0];for(c+="01010",e=4;e<8;e++)c+=v[R(r.charAt(e))][2]}else{for(f=A[R(r.charAt(0))],e=1;e<7;e++)c+=v[R(r.charAt(e))][R(f.charAt(e-1))];for(c+="01010",e=7;e<13;e++)c+=v[R(r.charAt(e))][2]}if(c+="101","ean13"===t)if(2===(a=s.substring(13,s.length)).length)for(c+="0000000000",i=parseInt(a,10)%4,e=0;e<2;e++)h=R(a.charAt(e)),l=R(p[R(i)][e]),c+=v[h][l];else if(5===a.length){for(c+="0000000000",u=!0,n=o=0,e=0;e<5;e++)u?o+=R(a.charAt(e)):n+=R(a.charAt(e)),u=!u;for(i=(9*n+3*o)%10,c+="1011",e=0;e<5;e++)h=R(a.charAt(e)),l=R(b[R(i)][e]),c+=v[h][l],e<4&&(c+="01")}return c}function N(r,t){var e,n="ean13"===t?12:7,o=r.substring(13,r.length),a=0,i=!0;for(e=(r=r.substring(0,n)).length-1;-1<e;e--)a+=(i?3:1)*R(r.charAt(e)),i=!i;return r+((10-a%10)%10).toString()+(o?" "+o:"")}var W=["101011","1101011","1001011","1100101","1011011","1101101","1001101","1010011","1101001","110101","101101"];var x=["101001101101","110100101011","101100101011","110110010101","101001101011","110100110101","101100110101","101001011011","110100101101","101100101101","110101001011","101101001011","110110100101","101011001011","110101100101","101101100101","101010011011","110101001101","101101001101","101011001101","110101010011","101101010011","110110101001","101011010011","110101101001","101101101001","101010110011","110101011001","101101011001","101011011001","110010101011","100110101011","110011010101","100101101011","110010110101","100110110101","100101011011","110010101101","100110101101","100100100101","100100101001","100101001001","101001001001","100101101101"];var C=["100010100","101001000","101000100","101000010","100101000","100100100","100100010","101010000","100010010","100001010","110101000","110100100","110100010","110010100","110010010","110001010","101101000","101100100","101100010","100110100","100011010","101011000","101001100","101000110","100101100","100010110","110110100","110110010","110101100","110100110","110010110","110011010","101101100","101100110","100110110","100111010","100101110","111010100","111010010","111001010","101101110","101110110","110101110","100100110","111011010","111010110","100110010","101011110"];var S=["11011001100","11001101100","11001100110","10010011000","10010001100","10001001100","10011001000","10011000100","10001100100","11001001000","11001000100","11000100100","10110011100","10011011100","10011001110","10111001100","10011101100","10011100110","11001110010","11001011100","11001001110","11011100100","11001110100","11101101110","11101001100","11100101100","11100100110","11101100100","11100110100","11100110010","11011011000","11011000110","11000110110","10100011000","10001011000","10001000110","10110001000","10001101000","10001100010","11010001000","11000101000","11000100010","10110111000","10110001110","10001101110","10111011000","10111000110","10001110110","11101110110","11010001110","11000101110","11011101000","11011100010","11011101110","11101011000","11101000110","11100010110","11101101000","11101100010","11100011010","11101111010","11001000010","11110001010","10100110000","10100001100","10010110000","10010000110","10000101100","10000100110","10110010000","10110000100","10011010000","10011000010","10000110100","10000110010","11000010010","11001010000","11110111010","11000010100","10001111010","10100111100","10010111100","10010011110","10111100100","10011110100","10011110010","11110100100","11110010100","11110010010","11011011110","11011110110","11110110110","10101111000","10100011110","10001011110","10111101000","10111100010","11110101000","11110100010","10111011110","10111101110","11101011110","11110101110","11010000100","11010010000","11010011100","11000111010"];var w=["101010011","101011001","101001011","110010101","101101001","110101001","100101011","100101101","100110101","110100101","101001101","101100101","1101011011","1101101011","1101101101","1011011011","1011001001","1010010011","1001001011","1010011001"];var I=["100100100100","100100100110","100100110100","100100110110","100110100100","100110100110","100110110100","100110110110","110100100100","110100100110"];function y(r,t){return"object"==typeof t?("mod10"===t.crc1?r=e(r):"mod11"===t.crc1&&(r=n(r)),"mod10"===t.crc2?r=e(r):"mod11"===t.crc2&&(r=n(r))):"boolean"==typeof t&&t&&(r=e(r)),r}function e(r){var t,e,n=r.length%2,o=0,a=0;for(t=0;t<r.length;t++)n?o=10*o+R(r.charAt(t)):a+=R(r.charAt(t)),n=!n;for(e=(2*o).toString(),t=0;t<e.length;t++)a+=R(e.charAt(t));return r+((10-a%10)%10).toString()}function n(r){var t,e=0,n=2;for(t=r.length-1;0<=t;t--)e+=n*R(r.charAt(t)),n=7===n?2:n+1;return r+((11-e%11)%11).toString()}var H=function(){var N=[10,12,14,16,18,20,22,24,26,32,36,40,44,48,52,64,72,80,88,96,104,120,132,144,8,8,12,12,16,16],W=[10,12,14,16,18,20,22,24,26,32,36,40,44,48,52,64,72,80,88,96,104,120,132,144,18,32,26,36,36,48],x=[3,5,8,12,18,22,30,36,44,62,86,114,144,174,204,280,368,456,576,696,816,1050,1304,1558,5,10,16,22,32,49],C=[5,7,10,12,14,18,20,24,28,36,42,48,56,68,84,112,144,192,224,272,336,408,496,620,7,11,14,18,24,28],S=[8,10,12,14,16,18,20,22,24,14,16,18,20,22,24,14,16,18,20,22,24,18,20,22,6,6,10,10,14,14],w=[8,10,12,14,16,18,20,22,24,14,16,18,20,22,24,14,16,18,20,22,24,18,20,22,16,14,24,16,16,22],I=[1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,4,4,4,4,4,4,6,6,6,1,1,1,1,1,1],y=[1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,4,4,4,4,4,4,6,6,6,1,2,1,2,2,2],H=[1,1,1,1,1,1,1,1,1,1,1,1,1,1,2,2,4,4,4,4,6,6,8,8,1,1,1,1,1,1],R=[-255,255,1,240,2,225,241,53,3,38,226,133,242,43,54,210,4,195,39,114,227,106,134,28,243,140,44,23,55,118,211,234,5,219,196,96,40,222,115,103,228,78,107,125,135,8,29,162,244,186,141,180,45,99,24,49,56,13,119,153,212,199,235,91,6,76,220,217,197,11,97,184,41,36,223,253,116,138,104,193,229,86,79,171,108,165,126,145,136,34,9,74,30,32,163,84,245,173,187,204,142,81,181,190,46,88,100,159,25,231,50,207,57,147,14,67,120,128,154,248,213,167,200,63,236,110,92,176,7,161,77,124,221,102,218,95,198,90,12,152,98,48,185,179,42,209,37,132,224,52,254,239,117,233,139,22,105,27,194,113,230,206,87,158,80,189,172,203,109,175,166,62,127,247,146,66,137,192,35,252,10,183,75,216,31,83,33,73,164,144,85,170,246,65,174,61,188,202,205,157,143,169,82,72,182,215,191,251,47,178,89,151,101,94,160,123,26,112,232,21,51,238,208,131,58,69,148,18,15,16,68,17,121,149,129,19,155,59,249,70,214,250,168,71,201,156,64,60,237,130,111,20,93,122,177,150],z=[1,2,4,8,16,32,64,128,45,90,180,69,138,57,114,228,229,231,227,235,251,219,155,27,54,108,216,157,23,46,92,184,93,186,89,178,73,146,9,18,36,72,144,13,26,52,104,208,141,55,110,220,149,7,14,28,56,112,224,237,247,195,171,123,246,193,175,115,230,225,239,243,203,187,91,182,65,130,41,82,164,101,202,185,95,190,81,162,105,210,137,63,126,252,213,135,35,70,140,53,106,212,133,39,78,156,21,42,84,168,125,250,217,159,19,38,76,152,29,58,116,232,253,215,131,43,86,172,117,234,249,223,147,11,22,44,88,176,77,154,25,50,100,200,189,87,174,113,226,233,255,211,139,59,118,236,245,199,163,107,214,129,47,94,188,85,170,121,242,201,191,83,166,97,194,169,127,254,209,143,51,102,204,181,71,142,49,98,196,165,103,206,177,79,158,17,34,68,136,61,122,244,197,167,99,198,161,111,222,145,15,30,60,120,240,205,183,67,134,33,66,132,37,74,148,5,10,20,40,80,160,109,218,153,31,62,124,248,221,151,3,6,12,24,48,96,192,173,119,238,241,207,179,75,150,1];function O(r,t){return r^t}function k(r){var t,e=[];for(t=0;t<8;t++)e[t]=r&128>>t?1:0;return e}function F(r,t,e,n,o,a,i){var c=X;c(r,t,e[0],n-2,o-2,a,i),c(r,t,e[1],n-2,o-1,a,i),c(r,t,e[2],n-1,o-2,a,i),c(r,t,e[3],n-1,o-1,a,i),c(r,t,e[4],n-1,o,a,i),c(r,t,e[5],n,o-2,a,i),c(r,t,e[6],n,o-1,a,i),c(r,t,e[7],n,o,a,i)}function X(r,t,e,n,o,a,i){n<0&&(n+=a,o+=4-(a+4)%8),o<0&&(o+=i,n+=4-(i+4)%8),1!==t[n][o]&&(r[n][o]=e,t[n][o]=1)}return function(r,t){var e,n=function(r){var t,e,n=[],o=0;for(t=0;t<r.length;t++)127<(e=r.charCodeAt(t))?(n[o]=235,e-=127,o++):48<=e&&e<=57&&t+1<r.length&&48<=r.charCodeAt(t+1)&&r.charCodeAt(t+1)<=57?(e=10*(e-48)+(r.charCodeAt(t+1)-48),e+=130,t++):e++,n[o]=e,o++;return n}(r),o=n.length,a=function(r,t){var e=0;if((r<1||1558<r)&&!t)return-1;if((r<1||49<r)&&t)return-1;for(t&&(e=24);x[e]<r;)e++;return e}(o,t),i=x[a],c=C[a],f=i+c,h=I[a],l=y[a],u=S[a],g=w[a],s=N[a]-2*h,d=W[a]-2*l,v=H[a],A=c/v,p=[],b=[],m=[];for(function(r,t,e){var n,o;if(!(e<=t))for(r[t]=129,o=t+1;o<e;o++)n=149*(o+1)%253+1,r[o]=(129+n)%254}(n,o,i),function(r,t,e,n,o){var a,i,c,f,h,l=0,u=r/o,g=[];for(c=0;c<o;c++){for(a=0;a<u;a++)g[a]=0;for(a=c;a<e;a+=o)for(l=O(n[a],g[u-1]),i=u-1;0<=i;i--)g[i]=l?(f=l,h=t[i],f&&h?z[(R[f]+R[h])%255]:0):0,0<i&&(g[i]=O(g[i-1],g[i]));for(i=e+c,a=u-1;0<=a;a--)n[i]=g[a],i+=o}}(c,function(r){var t,e,n,o,a=[];for(t=0;t<=r;t++)a[t]=1;for(t=1;t<=r;t++)for(e=t-1;0<=e;e--)a[e]=(n=a[e],o=t,n?o?z[(R[n]+o)%255]:n:0),0<e&&(a[e]=O(a[e],a[e-1]));return a}(A),i,n,v),e=0;e<f;e++)p[e]=k(n[e]);for(e=0;e<d;e++)b[e]=[],m[e]=[];return s*d%8==4&&(b[s-2][d-2]=1,b[s-1][d-1]=1,b[s-1][d-2]=0,b[s-2][d-1]=0,m[s-2][d-2]=1,m[s-1][d-1]=1,m[s-1][d-2]=1,m[s-2][d-1]=1),function(r,t,e,n,o,a){var i=0,c=4,f=0;do{for(c!==t||f?r<3&&c===t-2&&!f&&e%4?(y=o,H=a,R=n[i],z=t,O=e,k=void 0,(k=X)(y,H,R[0],z-3,0,z,O),k(y,H,R[1],z-2,0,z,O),k(y,H,R[2],z-1,0,z,O),k(y,H,R[3],0,O-4,z,O),k(y,H,R[4],0,O-3,z,O),k(y,H,R[5],0,O-2,z,O),k(y,H,R[6],0,O-1,z,O),k(y,H,R[7],1,O-1,z,O),i++):c!==t-2||f||e%8!=4?c===t+4&&2===f&&e%8==0&&(W=o,x=a,C=n[i],S=t,w=e,I=void 0,(I=X)(W,x,C[0],S-1,0,S,w),I(W,x,C[1],S-1,w-1,S,w),I(W,x,C[2],0,w-3,S,w),I(W,x,C[3],0,w-2,S,w),I(W,x,C[4],0,w-1,S,w),I(W,x,C[5],1,w-3,S,w),I(W,x,C[6],1,w-2,S,w),I(W,x,C[7],1,w-1,S,w),i++):(v=o,A=a,p=n[i],b=t,m=e,N=void 0,(N=X)(v,A,p[0],b-3,0,b,m),N(v,A,p[1],b-2,0,b,m),N(v,A,p[2],b-1,0,b,m),N(v,A,p[3],0,m-2,b,m),N(v,A,p[4],0,m-1,b,m),N(v,A,p[5],1,m-1,b,m),N(v,A,p[6],2,m-1,b,m),N(v,A,p[7],3,m-1,b,m),i++):(h=o,l=a,u=n[i],g=t,s=e,d=void 0,(d=X)(h,l,u[0],g-1,0,g,s),d(h,l,u[1],g-1,1,g,s),d(h,l,u[2],g-1,2,g,s),d(h,l,u[3],0,s-2,g,s),d(h,l,u[4],0,s-1,g,s),d(h,l,u[5],1,s-1,g,s),d(h,l,u[6],2,s-1,g,s),d(h,l,u[7],3,s-1,g,s),i++);c<t&&0<=f&&1!==a[c][f]&&(F(o,a,n[i],c,f,t,e),i++),f+=2,0<=(c-=2)&&f<e;);for(c+=1,f+=3;0<=c&&f<e&&1!==a[c][f]&&(F(o,a,n[i],c,f,t,e),i++),f-=2,(c+=2)<t&&0<=f;);c+=3,f+=1}while(c<t||f<e);var h,l,u,g,s,d;var v,A,p,b,m,N;var W,x,C,S,w,I;var y,H,R,z,O,k}(0,s,d,p,b,m),b=function(r,t,e,n,o){var a,i,c=(n+2)*t,f=(o+2)*e,h=[];for(h[0]=[],i=0;i<f+2;i++)h[0][i]=0;for(a=0;a<c;a++)for(h[a+1]=[],h[a+1][0]=0,i=h[a+1][f+1]=0;i<f;i++)h[a+1][i+1]=a%(n+2)==0?i%2?0:1:a%(n+2)===n+1?1:i%(o+2)===o+1?a%2?1:0:i%(o+2)==0?1:(h[a+1][i+1]=0,r[a-1-2*parseInt(a/(n+2),10)][i-1-2*parseInt(i/(o+2),10)]);for(h[c+1]=[],i=0;i<f+2;i++)h[c+1][i]=0;return h}(b,h,l,u,g)}}();function R(r){var t=typeof r;return"string"===t?(r=r.replace(/[^0-9-.]/g,""),r=parseInt(1*r,10),isNaN(r)||!isFinite(r)?0:r):"number"===t&&isFinite(r)?Math.floor(r):0}function z(r,t){var e,n="";for(e=0;e<t;e++)n+=String.fromCharCode(255&r),r>>=8;return n}function O(r,t,e){return String.fromCharCode(e)+String.fromCharCode(t)+String.fromCharCode(r)}function k(r){var t=parseInt("0x"+r.substr(1),16),e=255&t;return O((t>>=8)>>8,255&t,e)}function F(r){return r.match(/#[0-91-F]/gi)}function X(r){for(var t,e,n,o,a,i,c,f="",h="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",l=0;l<r.length;)o=(t=r.charCodeAt(l++))>>2,a=(3&t)<<4|(e=r.charCodeAt(l++))>>4,i=(15&e)<<2|(n=r.charCodeAt(l++))>>6,c=63&n,isNaN(e)?i=c=64:isNaN(n)&&(c=64),f+=h.charAt(o)+h.charAt(a)+h.charAt(i)+h.charAt(c);return f}function f(r){var t,e=[];for(e[0]=[],t=0;t<r.length;t++)e[0][t]=parseInt(r.charAt(t),10);return e}function Y(r,t){return r.css("padding","0").css("overflow","auto").css("width",t+"px").html(""),r}function i(r,t,e,n,o,a){var i,c,f,h,l,u,g,s,d,v=e.length,A=e[0].length,p=F(t.bgColor)?k(t.bgColor):O(255,255,255),b=F(t.color)?k(t.color):O(0,0,0),m="",N="",W="";for(i=0;i<o;i++)m+=p,N+=b;for(u=(o*A+(l=(4-o*A*3%4)%4))*a*v,i=0;i<l;i++)W+="\0";for(g="BM"+z(54+u,4)+"\0\0\0\0"+z(54,4)+z(40,4)+z(o*A,4)+z(a*v,4)+z(1,2)+z(24,2)+"\0\0\0\0"+z(u,4)+z(2835,4)+z(2835,4)+z(0,4)+z(0,4),c=v-1;0<=c;c--){for(s="",f=0;f<A;f++)s+=e[c][f]?N:m;for(s+=W,h=0;h<a;h++)g+=s}(d=document.createElement("object")).setAttribute("type","image/bmp"),d.setAttribute("data","data:image/bmp;base64,"+X(g)),Y(r,o*A+l).append(d)}function c(r,t,e,n,o,a){var i,c,f,h,l=e.length,u=e[0].length,g="",s='<div style="float: left; font-size: 0; background-color: '+t.bgColor+"; height: "+a+'px; width: &Wpx"></div>',d='<div style="float: left; font-size: 0; width:0; border-left: &Wpx solid '+t.color+"; height: "+a+'px;"></div>';for(c=0;c<l;c++){for(f=0,h=e[c][0],i=0;i<u;i++)h===e[c][i]?f++:(g+=(h?d:s).replace("&W",f*o),h=e[c][i],f=1);0<f&&(g+=(h?d:s).replace("&W",f*o))}t.showHRI&&(g+='<div style="clear:both; width: 100%; background-color: '+t.bgColor+"; color: "+t.color+"; text-align: center; font-size: "+t.fontSize+"px; margin-top: "+t.marginHRI+'px;">'+n+"</div>"),Y(r,o*u).html(g)}function l(r,t,e,n,o,a){var i,c,f,h,l,u,g,s,d=e.length,v=e[0].length,A=o*v,p=a*d;for(t.showHRI&&(f=R(t.fontSize),p+=R(t.marginHRI)+f),h='<svg xmlns="http://www.w3.org/2000/svg" version="1.1" width="'+A+'" height="'+p+'">',h+='<rect width="'+A+'" height="'+p+'" x="0" y="0" fill="'+t.bgColor+'" />',l='<rect width="&W" height="'+a+'" x="&X" y="&Y" fill="'+t.color+'" />',c=0;c<d;c++){for(u=0,g=e[c][0],i=0;i<v;i++)g===e[c][i]?u++:(g&&(h+=l.replace("&W",u*o).replace("&X",(i-u)*o).replace("&Y",c*a)),g=e[c][i],u=1);u&&g&&(h+=l.replace("&W",u*o).replace("&X",(v-u)*o).replace("&Y",c*a))}t.showHRI&&(h+='<g transform="translate('+Math.floor(A/2)+' 0)">',h+='<text y="'+(p-Math.floor(f/2))+'" text-anchor="middle" style="font-family: Arial; font-size: '+f+'px;" fill="'+t.color+'">'+n+"</text>",h+="</g>"),h+="</svg>",(s=document.createElement("img")).setAttribute("src","data:image/svg+xml;base64,"+X(h)),Y(r,A).append(s)}function u(r,t,e,n,o,a,i,c){var f,h,l,u,g,s,d=r.get(0),v=e.length,A=e[0].length;if(d&&d.getContext){for((l=d.getContext("2d")).lineWidth=1,l.lineCap="butt",l.fillStyle=t.bgColor,l.fillRect(o,a,A*i,v*c),l.fillStyle=t.color,h=0;h<v;h++){for(u=0,g=e[h][0],f=0;f<A;f++)g===e[h][f]?u++:(g&&l.fillRect(o+(f-u)*i,a+h*c,i*u,c),g=e[h][f],u=1);u&&g&&l.fillRect(o+(A-u)*i,a+h*c,i*u,c)}t.showHRI&&(s=l.measureText(n),l.fillText(n,o+Math.floor((A*i-s.width)/2),a+v*c+t.fontSize+t.marginHRI))}}var M={bmp:function(r,t,e,n){var o=R(t.barWidth),a=R(t.barHeight);i(r,t,f(e),0,o,a)},bmp2:function(r,t,e,n){var o=R(t.moduleSize);i(r,t,e,0,o,o)},css:function(r,t,e,n){var o=R(t.barWidth),a=R(t.barHeight);c(r,t,f(e),n,o,a)},css2:function(r,t,e,n){var o=R(t.moduleSize);c(r,t,e,n,o,o)},svg:function(r,t,e,n){var o=R(t.barWidth),a=R(t.barHeight);l(r,t,f(e),n,o,a)},svg2:function(r,t,e,n){var o=R(t.moduleSize);l(r,t,e,n,o,o)},canvas:function(r,t,e,n){var o=R(t.barWidth),a=R(t.barHeight),i=R(t.posX),c=R(t.posY);u(r,t,f(e),n,i,c,o,a)},canvas2:function(r,t,e,n){var o=R(t.moduleSize);u(r,t,e,n,R(t.posX),R(t.posY),o,o)}};g.fn.barcode=function(r,t,e){var n,o,a,i,c,f,h="",l="",u=!1;if(n=(r=g.extend({crc:!0,rect:!1},"object"==typeof r?r:{code:r})).code,o=r.crc,a=r.rect,n){switch(e=g.extend(!0,s,e),t){case"std25":case"int25":h=d(n,o,t),l=d(n,o,t);break;case"ean8":case"ean13":h=m(n,t),l=N(n,t);break;case"upc":(f=n).length<12&&(f="0"+f),h=m(f,"ean13"),(c=n).length<12&&(c="0"+c),l=N(c,"ean13").substr(1);break;case"code11":h=function(r){var t,e,n,o,a,i,c,f,h="0123456789-",l="10110010";for(n=0;n<r.length;n++){if((o=h.indexOf(r.charAt(n)))<0)return"";l+=W[o]+"0"}for(c=1,f=i=a=0,n=r.length-1;0<=n;n--)c=10===c?1:c+1,i+=(a=10===a?1:a+1)*(o=h.indexOf(r.charAt(n))),f+=c*o;return e=(f+=t=i%11)%11,l+=W[t]+"0",10<=r.length&&(l+=W[e]+"0"),l+="1011001"}(n),l=n;break;case"code39":h=function(r){var t,e,n="";if(0<=r.indexOf("*"))return"";for(r=("*"+r+"*").toUpperCase(),t=0;t<r.length;t++){if((e="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*".indexOf(r.charAt(t)))<0)return"";0<t&&(n+="0"),n+=x[e]}return n}(n),l=n;break;case"code93":h=function(r,t){var e,n,o,a,i,c,f,h="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%____*",l="";if(0<=r.indexOf("*"))return"";for(r=r.toUpperCase(),l+=C[47],n=0;n<r.length;n++){if(e=r.charAt(n),o=h.indexOf(e),"_"===e||o<0)return"";l+=C[o]}if(t){for(c=1,f=i=a=0,n=r.length-1;0<=n;n--)c=15===c?1:c+1,i+=(a=20===a?1:a+1)*(o=h.indexOf(r.charAt(n))),f+=c*o;l+=C[e=i%47],l+=C[(f+=e)%47]}return l+=C[47],l+="1"}(n,o),l=n;break;case"code128":h=function(r){var t,e,n,o,a,i=" !\"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~",c=0,f=0,h=0;for(t=0;t<r.length;t++)if(-1===i.indexOf(r.charAt(t)))return"";for(n=1<r.length,t=0;t<3&&t<r.length;t++)e=r.charAt(t),n=n&&"0"<=e&&e<="9";for(o=S[a=n?105:104],t=0;t<r.length;){if(n)(t===r.length||r.charAt(t)<"0"||"9"<r.charAt(t)||r.charAt(t+1)<"0"||"9"<r.charAt(t+1))&&(n=!1,o+=S[100],a+=100*++c);else{for(f=0;t+f<r.length&&"0"<=r.charAt(t+f)&&r.charAt(t+f)<="9";)f++;(n=5<f||t+f-1===r.length&&3<f)&&(o+=S[99],a+=99*++c)}n?(h=R(r.charAt(t)+r.charAt(t+1)),t+=2):(h=i.indexOf(r.charAt(t)),t+=1),o+=S[h],a+=++c*h}return o+=S[a%103],o+=S[106],o+="11"}(n),l=n;break;case"codabar":h=function(r){var t,e,n;for(n=w[16]+"0",t=0;t<r.length;t++){if((e="0123456789-$:/.+".indexOf(r.charAt(t)))<0)return"";n+=w[e]+"0"}return n+=w[16]}(n),l=n;break;case"msi":h=function(r){var t,e=0,n="110";for(r=y(r,!1),t=0;t<r.length;t++){if((e="0123456789".indexOf(r.charAt(t)))<0)return"";n+=I[e]}return n+="1001"}(n),l=y(n,o);break;case"datamatrix":h=H(n,a),l=n,u=!0}h.length&&(!u&&e.addQuietZone&&(h="0000000000"+h+"0000000000"),(i=M[e.output+(u?"2":"")])&&this.each(function(){i(g(this),e,h,l)}))}return this}}(jQuery);

/*! jquery-qrcode v0.17.0 - https://larsjung.de/jquery-qrcode/ */
!function(t,r){"object"==typeof exports&&"object"==typeof module?module.exports=r():"function"==typeof define&&define.amd?define("jquery-qrcode",[],r):"object"==typeof exports?exports["jquery-qrcode"]=r():t["jquery-qrcode"]=r()}("undefined"!=typeof self?self:this,function(){return function(e){var n={};function o(t){if(n[t])return n[t].exports;var r=n[t]={i:t,l:!1,exports:{}};return e[t].call(r.exports,r,r.exports,o),r.l=!0,r.exports}return o.m=e,o.c=n,o.d=function(t,r,e){o.o(t,r)||Object.defineProperty(t,r,{enumerable:!0,get:e})},o.r=function(t){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(t,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(t,"__esModule",{value:!0})},o.t=function(r,t){if(1&t&&(r=o(r)),8&t)return r;if(4&t&&"object"==typeof r&&r&&r.__esModule)return r;var e=Object.create(null);if(o.r(e),Object.defineProperty(e,"default",{enumerable:!0,value:r}),2&t&&"string"!=typeof r)for(var n in r)o.d(e,n,function(t){return r[t]}.bind(null,n));return e},o.n=function(t){var r=t&&t.__esModule?function(){return t.default}:function(){return t};return o.d(r,"a",r),r},o.o=function(t,r){return Object.prototype.hasOwnProperty.call(t,r)},o.p="",o(o.s=0)}([function(v,t,p){(function(t){function c(t){return t&&"string"==typeof t.tagName&&"IMG"===t.tagName.toUpperCase()}function a(t,r,e,n){var o={},i=p(2);i.stringToBytes=i.stringToBytesFuncs["UTF-8"];var a=i(e,r);a.addData(t),a.make(),n=n||0;var u=a.getModuleCount(),s=u+2*n;return o.text=t,o.level=r,o.version=e,o.module_count=s,o.is_dark=function(t,r){return r-=n,0<=(t-=n)&&t<u&&0<=r&&r<u&&a.isDark(t,r)},o.add_blank=function(a,u,f,c){var l=o.is_dark,g=1/s;o.is_dark=function(t,r){var e=r*g,n=t*g,o=e+g,i=n+g;return l(t,r)&&(o<a||f<e||i<u||c<n)}},o}function h(t,r,e,n,o){e=Math.max(1,e||1),n=Math.min(40,n||40);for(var i=e;i<=n;i+=1)try{return a(t,r,i,o)}catch(t){}}function i(t,r,e){c(e.background)?r.drawImage(e.background,0,0,e.size,e.size):e.background&&(r.fillStyle=e.background,r.fillRect(e.left,e.top,e.size,e.size));var n=e.mode;1===n||2===n?function(t,r,e){var n=e.size,o="bold "+e.mSize*n+"px "+e.fontname,i=d("<canvas/>")[0].getContext("2d");i.font=o;var a=i.measureText(e.label).width,u=e.mSize,f=a/n,c=(1-f)*e.mPosX,l=(1-u)*e.mPosY,g=c+f,s=l+u;1===e.mode?t.add_blank(0,l-.01,n,s+.01):t.add_blank(c-.01,l-.01,.01+g,s+.01),r.fillStyle=e.fontcolor,r.font=o,r.fillText(e.label,c*n,l*n+.75*e.mSize*n)}(t,r,e):!c(e.image)||3!==n&&4!==n||function(t,r,e){var n=e.size,o=e.image.naturalWidth||1,i=e.image.naturalHeight||1,a=e.mSize,u=a*o/i,f=(1-u)*e.mPosX,c=(1-a)*e.mPosY,l=f+u,g=c+a;3===e.mode?t.add_blank(0,c-.01,n,g+.01):t.add_blank(f-.01,c-.01,.01+l,g+.01),r.drawImage(e.image,f*n,c*n,u*n,a*n)}(t,r,e)}function l(t,r,e,n,o,i,a,u){t.is_dark(a,u)&&r.rect(n,o,i,i)}function g(t,r,e,n,o,i,a,u){var f=t.is_dark,c=n+i,l=o+i,g=e.radius*i,s=a-1,h=a+1,d=u-1,v=u+1,p=f(a,u),w=f(s,d),y=f(s,u),m=f(s,v),b=f(a,v),k=f(h,v),C=f(h,u),B=f(h,d),x=f(a,d);p?function(t,r,e,n,o,i,a,u,f,c){a?t.moveTo(r+i,e):t.moveTo(r,e),u?(t.lineTo(n-i,e),t.arcTo(n,e,n,o,i)):t.lineTo(n,e),f?(t.lineTo(n,o-i),t.arcTo(n,o,r,o,i)):t.lineTo(n,o),c?(t.lineTo(r+i,o),t.arcTo(r,o,r,e,i)):t.lineTo(r,o),a?(t.lineTo(r,e+i),t.arcTo(r,e,n,e,i)):t.lineTo(r,e)}(r,n,o,c,l,g,!y&&!x,!y&&!b,!C&&!b,!C&&!x):function(t,r,e,n,o,i,a,u,f,c){a&&(t.moveTo(r+i,e),t.lineTo(r,e),t.lineTo(r,e+i),t.arcTo(r,e,r+i,e,i)),u&&(t.moveTo(n-i,e),t.lineTo(n,e),t.lineTo(n,e+i),t.arcTo(n,e,n-i,e,i)),f&&(t.moveTo(n-i,o),t.lineTo(n,o),t.lineTo(n,o-i),t.arcTo(n,o,n-i,o,i)),c&&(t.moveTo(r+i,o),t.lineTo(r,o),t.lineTo(r,o-i),t.arcTo(r,o,r+i,o,i))}(r,n,o,c,l,g,y&&x&&w,y&&b&&m,C&&b&&k,C&&x&&B)}function n(t,r){var e=h(r.text,r.ecLevel,r.minVersion,r.maxVersion,r.quiet);if(!e)return null;var n=d(t).data("qrcode",e),o=n[0].getContext("2d");return i(e,o,r),function(t,r,e){var n,o,i=t.module_count,a=e.size/i,u=l;for(0<e.radius&&e.radius<=.5&&(u=g),r.beginPath(),n=0;n<i;n+=1)for(o=0;o<i;o+=1)u(t,r,e,e.left+o*a,e.top+n*a,a,n,o);if(c(e.fill)){r.strokeStyle="rgba(0,0,0,0.5)",r.lineWidth=2,r.stroke();var f=r.globalCompositeOperation;r.globalCompositeOperation="destination-out",r.fill(),r.globalCompositeOperation=f,r.clip(),r.drawImage(e.fill,0,0,e.size,e.size),r.restore()}else r.fillStyle=e.fill,r.fill()}(e,o,r),n}function r(t){var r=d("<canvas/>").attr("width",t.size).attr("height",t.size);return n(r,t)}function o(t){return f&&"canvas"===t.render?r(t):f&&"image"===t.render?function(t){return d("<img/>").attr("src",r(t)[0].toDataURL("image/png"))}(t):function(t){var r=h(t.text,t.ecLevel,t.minVersion,t.maxVersion,t.quiet);if(!r)return null;var e,n,o=t.size,i=t.background,a=Math.floor,u=r.module_count,f=a(o/u),c=a(.5*(o-f*u)),l={position:"relative",left:0,top:0,padding:0,margin:0,width:o,height:o},g={position:"absolute",padding:0,margin:0,width:f,height:f,"background-color":t.fill},s=d("<div/>").data("qrcode",r).css(l);for(i&&s.css("background-color",i),e=0;e<u;e+=1)for(n=0;n<u;n+=1)r.is_dark(e,n)&&d("<div/>").css(g).css({left:c+n*f,top:c+e*f}).appendTo(s);return s}(t)}var e,u=t.window,d=u.jQuery,f=!(!(e=u.document.createElement("canvas")).getContext||!e.getContext("2d")),s={render:"canvas",minVersion:1,maxVersion:40,ecLevel:"L",left:0,top:0,size:200,fill:"#000",background:"#fff",text:"no text",radius:0,quiet:0,mode:0,mSize:.1,mPosX:.5,mPosY:.5,label:"no label",fontname:"sans",fontcolor:"#000",image:null};d.fn.qrcode=v.exports=function(t){var e=d.extend({},s,t);return this.each(function(t,r){"canvas"===r.nodeName.toLowerCase()?n(r,e):d(r).append(o(e))})}}).call(this,p(1))},function(t,r){var e;e=function(){return this}();try{e=e||new Function("return this")()}catch(t){"object"==typeof window&&(e=window)}t.exports=e},function(t,r,e){var n,o,i,a=function(){function i(t,r){function a(t,r){l=function(t){for(var r=new Array(t),e=0;e<t;e+=1){r[e]=new Array(t);for(var n=0;n<t;n+=1)r[e][n]=null}return r}(g=4*u+17),e(0,0),e(g-7,0),e(0,g-7),i(),o(),d(t,r),7<=u&&s(t),null==n&&(n=p(u,f,c)),v(n,r)}var u=t,f=w[r],l=null,g=0,n=null,c=[],h={},e=function(t,r){for(var e=-1;e<=7;e+=1)if(!(t+e<=-1||g<=t+e))for(var n=-1;n<=7;n+=1)r+n<=-1||g<=r+n||(l[t+e][r+n]=0<=e&&e<=6&&(0==n||6==n)||0<=n&&n<=6&&(0==e||6==e)||2<=e&&e<=4&&2<=n&&n<=4)},o=function(){for(var t=8;t<g-8;t+=1)null==l[t][6]&&(l[t][6]=t%2==0);for(var r=8;r<g-8;r+=1)null==l[6][r]&&(l[6][r]=r%2==0)},i=function(){for(var t=y.getPatternPosition(u),r=0;r<t.length;r+=1)for(var e=0;e<t.length;e+=1){var n=t[r],o=t[e];if(null==l[n][o])for(var i=-2;i<=2;i+=1)for(var a=-2;a<=2;a+=1)l[n+i][o+a]=-2==i||2==i||-2==a||2==a||0==i&&0==a}},s=function(t){for(var r=y.getBCHTypeNumber(u),e=0;e<18;e+=1){var n=!t&&1==(r>>e&1);l[Math.floor(e/3)][e%3+g-8-3]=n}for(e=0;e<18;e+=1){n=!t&&1==(r>>e&1);l[e%3+g-8-3][Math.floor(e/3)]=n}},d=function(t,r){for(var e=f<<3|r,n=y.getBCHTypeInfo(e),o=0;o<15;o+=1){var i=!t&&1==(n>>o&1);o<6?l[o][8]=i:o<8?l[o+1][8]=i:l[g-15+o][8]=i}for(o=0;o<15;o+=1){i=!t&&1==(n>>o&1);o<8?l[8][g-o-1]=i:o<9?l[8][15-o-1+1]=i:l[8][15-o-1]=i}l[g-8][8]=!t},v=function(t,r){for(var e=-1,n=g-1,o=7,i=0,a=y.getMaskFunction(r),u=g-1;0<u;u-=2)for(6==u&&(u-=1);;){for(var f=0;f<2;f+=1)if(null==l[n][u-f]){var c=!1;i<t.length&&(c=1==(t[i]>>>o&1)),a(n,u-f)&&(c=!c),l[n][u-f]=c,-1==(o-=1)&&(i+=1,o=7)}if((n+=e)<0||g<=n){n-=e,e=-e;break}}},p=function(t,r,e){for(var n=C.getRSBlocks(t,r),o=B(),i=0;i<e.length;i+=1){var a=e[i];o.put(a.getMode(),4),o.put(a.getLength(),y.getLengthInBits(a.getMode(),t)),a.write(o)}var u=0;for(i=0;i<n.length;i+=1)u+=n[i].dataCount;if(o.getLengthInBits()>8*u)throw"code length overflow. ("+o.getLengthInBits()+">"+8*u+")";for(o.getLengthInBits()+4<=8*u&&o.put(0,4);o.getLengthInBits()%8!=0;)o.putBit(!1);for(;!(o.getLengthInBits()>=8*u||(o.put(236,8),o.getLengthInBits()>=8*u));)o.put(17,8);return function(t,r){for(var e=0,n=0,o=0,i=new Array(r.length),a=new Array(r.length),u=0;u<r.length;u+=1){var f=r[u].dataCount,c=r[u].totalCount-f;n=Math.max(n,f),o=Math.max(o,c),i[u]=new Array(f);for(var l=0;l<i[u].length;l+=1)i[u][l]=255&t.getBuffer()[l+e];e+=f;var g=y.getErrorCorrectPolynomial(c),s=m(i[u],g.getLength()-1).mod(g);for(a[u]=new Array(g.getLength()-1),l=0;l<a[u].length;l+=1){var h=l+s.getLength()-a[u].length;a[u][l]=0<=h?s.getAt(h):0}}var d=0;for(l=0;l<r.length;l+=1)d+=r[l].totalCount;var v=new Array(d),p=0;for(l=0;l<n;l+=1)for(u=0;u<r.length;u+=1)l<i[u].length&&(v[p]=i[u][l],p+=1);for(l=0;l<o;l+=1)for(u=0;u<r.length;u+=1)l<a[u].length&&(v[p]=a[u][l],p+=1);return v}(o,n)};return h.addData=function(t,r){var e=null;switch(r=r||"Byte"){case"Numeric":e=x(t);break;case"Alphanumeric":e=T(t);break;case"Byte":e=M(t);break;case"Kanji":e=A(t);break;default:throw"mode:"+r}c.push(e),n=null},h.isDark=function(t,r){if(t<0||g<=t||r<0||g<=r)throw t+","+r;return l[t][r]},h.getModuleCount=function(){return g},h.make=function(){if(u<1){for(var t=1;t<40;t++){for(var r=C.getRSBlocks(t,f),e=B(),n=0;n<c.length;n++){var o=c[n];e.put(o.getMode(),4),e.put(o.getLength(),y.getLengthInBits(o.getMode(),t)),o.write(e)}var i=0;for(n=0;n<r.length;n++)i+=r[n].dataCount;if(e.getLengthInBits()<=8*i)break}u=t}a(!1,function(){for(var t=0,r=0,e=0;e<8;e+=1){a(!0,e);var n=y.getLostPoint(h);(0==e||n<t)&&(t=n,r=e)}return r}())},h.createTableTag=function(t,r){t=t||2;var e="";e+='<table style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: "+(r=void 0===r?4*t:r)+"px;",e+='">',e+="<tbody>";for(var n=0;n<h.getModuleCount();n+=1){e+="<tr>";for(var o=0;o<h.getModuleCount();o+=1)e+='<td style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: 0px;",e+=" width: "+t+"px;",e+=" height: "+t+"px;",e+=" background-color: ",e+=h.isDark(n,o)?"#000000":"#ffffff",e+=";",e+='"/>';e+="</tr>"}return e+="</tbody>",e+="</table>"},h.createSvgTag=function(t,r){var e={};"object"==typeof t&&(t=(e=t).cellSize,r=e.margin),t=t||2,r=void 0===r?4*t:r;var n,o,i,a,u=h.getModuleCount()*t+2*r,f="";for(a="l"+t+",0 0,"+t+" -"+t+",0 0,-"+t+"z ",f+='<svg version="1.1" xmlns="http://www.w3.org/2000/svg"',f+=e.scalable?"":' width="'+u+'px" height="'+u+'px"',f+=' viewBox="0 0 '+u+" "+u+'" ',f+=' preserveAspectRatio="xMinYMin meet">',f+='<rect width="100%" height="100%" fill="white" cx="0" cy="0"/>',f+='<path d="',o=0;o<h.getModuleCount();o+=1)for(i=o*t+r,n=0;n<h.getModuleCount();n+=1)h.isDark(o,n)&&(f+="M"+(n*t+r)+","+i+a);return f+='" stroke="transparent" fill="black"/>',f+="</svg>"},h.createDataURL=function(o,t){o=o||2,t=void 0===t?4*o:t;var r=h.getModuleCount()*o+2*t,i=t,a=r-t;return L(r,r,function(t,r){if(i<=t&&t<a&&i<=r&&r<a){var e=Math.floor((t-i)/o),n=Math.floor((r-i)/o);return h.isDark(n,e)?0:1}return 1})},h.createImgTag=function(t,r,e){t=t||2,r=void 0===r?4*t:r;var n=h.getModuleCount()*t+2*r,o="";return o+="<img",o+=' src="',o+=h.createDataURL(t,r),o+='"',o+=' width="',o+=n,o+='"',o+=' height="',o+=n,o+='"',e&&(o+=' alt="',o+=e,o+='"'),o+="/>"},h.createASCII=function(t,r){if((t=t||1)<2)return function(t){t=void 0===t?2:t;var r,e,n,o,i,a=1*h.getModuleCount()+2*t,u=t,f=a-t,c={"██":"█","█ ":"▀"," █":"▄","  ":" "},l={"██":"▀","█ ":"▀"," █":" ","  ":" "},g="";for(r=0;r<a;r+=2){for(n=Math.floor((r-u)/1),o=Math.floor((r+1-u)/1),e=0;e<a;e+=1)i="█",u<=e&&e<f&&u<=r&&r<f&&h.isDark(n,Math.floor((e-u)/1))&&(i=" "),u<=e&&e<f&&u<=r+1&&r+1<f&&h.isDark(o,Math.floor((e-u)/1))?i+=" ":i+="█",g+=t<1&&f<=r+1?l[i]:c[i];g+="\n"}return a%2&&0<t?g.substring(0,g.length-a-1)+Array(1+a).join("▀"):g.substring(0,g.length-1)}(r);t-=1,r=void 0===r?2*t:r;var e,n,o,i,a=h.getModuleCount()*t+2*r,u=r,f=a-r,c=Array(t+1).join("██"),l=Array(t+1).join("  "),g="",s="";for(e=0;e<a;e+=1){for(o=Math.floor((e-u)/t),s="",n=0;n<a;n+=1)i=1,u<=n&&n<f&&u<=e&&e<f&&h.isDark(o,Math.floor((n-u)/t))&&(i=0),s+=i?c:l;for(o=0;o<t;o+=1)g+=s+"\n"}return g.substring(0,g.length-1)},h.renderTo2dContext=function(t,r){r=r||2;for(var e=h.getModuleCount(),n=0;n<e;n++)for(var o=0;o<e;o++)t.fillStyle=h.isDark(n,o)?"black":"white",t.fillRect(n*r,o*r,r,r)},h}i.stringToBytes=(i.stringToBytesFuncs={default:function(t){for(var r=[],e=0;e<t.length;e+=1){var n=t.charCodeAt(e);r.push(255&n)}return r}}).default,i.createStringToBytes=function(u,f){var i=function(){function t(){var t=r.read();if(-1==t)throw"eof";return t}for(var r=S(u),e=0,n={};;){var o=r.read();if(-1==o)break;var i=t(),a=t()<<8|t();n[String.fromCharCode(o<<8|i)]=a,e+=1}if(e!=f)throw e+" != "+f;return n}(),a="?".charCodeAt(0);return function(t){for(var r=[],e=0;e<t.length;e+=1){var n=t.charCodeAt(e);if(n<128)r.push(n);else{var o=i[t.charAt(e)];"number"==typeof o?(255&o)==o?r.push(o):(r.push(o>>>8),r.push(255&o)):r.push(a)}}return r}};var a=1,u=2,o=4,f=8,w={L:1,M:0,Q:3,H:2},n=0,c=1,l=2,g=3,s=4,h=5,d=6,v=7,y=function(){function e(t){for(var r=0;0!=t;)r+=1,t>>>=1;return r}var r=[[],[6,18],[6,22],[6,26],[6,30],[6,34],[6,22,38],[6,24,42],[6,26,46],[6,28,50],[6,30,54],[6,32,58],[6,34,62],[6,26,46,66],[6,26,48,70],[6,26,50,74],[6,30,54,78],[6,30,56,82],[6,30,58,86],[6,34,62,90],[6,28,50,72,94],[6,26,50,74,98],[6,30,54,78,102],[6,28,54,80,106],[6,32,58,84,110],[6,30,58,86,114],[6,34,62,90,118],[6,26,50,74,98,122],[6,30,54,78,102,126],[6,26,52,78,104,130],[6,30,56,82,108,134],[6,34,60,86,112,138],[6,30,58,86,114,142],[6,34,62,90,118,146],[6,30,54,78,102,126,150],[6,24,50,76,102,128,154],[6,28,54,80,106,132,158],[6,32,58,84,110,136,162],[6,26,54,82,110,138,166],[6,30,58,86,114,142,170]],t={};return t.getBCHTypeInfo=function(t){for(var r=t<<10;0<=e(r)-e(1335);)r^=1335<<e(r)-e(1335);return 21522^(t<<10|r)},t.getBCHTypeNumber=function(t){for(var r=t<<12;0<=e(r)-e(7973);)r^=7973<<e(r)-e(7973);return t<<12|r},t.getPatternPosition=function(t){return r[t-1]},t.getMaskFunction=function(t){switch(t){case n:return function(t,r){return(t+r)%2==0};case c:return function(t,r){return t%2==0};case l:return function(t,r){return r%3==0};case g:return function(t,r){return(t+r)%3==0};case s:return function(t,r){return(Math.floor(t/2)+Math.floor(r/3))%2==0};case h:return function(t,r){return t*r%2+t*r%3==0};case d:return function(t,r){return(t*r%2+t*r%3)%2==0};case v:return function(t,r){return(t*r%3+(t+r)%2)%2==0};default:throw"bad maskPattern:"+t}},t.getErrorCorrectPolynomial=function(t){for(var r=m([1],0),e=0;e<t;e+=1)r=r.multiply(m([1,p.gexp(e)],0));return r},t.getLengthInBits=function(t,r){if(1<=r&&r<10)switch(t){case a:return 10;case u:return 9;case o:case f:return 8;default:throw"mode:"+t}else if(r<27)switch(t){case a:return 12;case u:return 11;case o:return 16;case f:return 10;default:throw"mode:"+t}else{if(!(r<41))throw"type:"+r;switch(t){case a:return 14;case u:return 13;case o:return 16;case f:return 12;default:throw"mode:"+t}}},t.getLostPoint=function(t){for(var r=t.getModuleCount(),e=0,n=0;n<r;n+=1)for(var o=0;o<r;o+=1){for(var i=0,a=t.isDark(n,o),u=-1;u<=1;u+=1)if(!(n+u<0||r<=n+u))for(var f=-1;f<=1;f+=1)o+f<0||r<=o+f||0==u&&0==f||a==t.isDark(n+u,o+f)&&(i+=1);5<i&&(e+=3+i-5)}for(n=0;n<r-1;n+=1)for(o=0;o<r-1;o+=1){var c=0;t.isDark(n,o)&&(c+=1),t.isDark(n+1,o)&&(c+=1),t.isDark(n,o+1)&&(c+=1),t.isDark(n+1,o+1)&&(c+=1),0!=c&&4!=c||(e+=3)}for(n=0;n<r;n+=1)for(o=0;o<r-6;o+=1)t.isDark(n,o)&&!t.isDark(n,o+1)&&t.isDark(n,o+2)&&t.isDark(n,o+3)&&t.isDark(n,o+4)&&!t.isDark(n,o+5)&&t.isDark(n,o+6)&&(e+=40);for(o=0;o<r;o+=1)for(n=0;n<r-6;n+=1)t.isDark(n,o)&&!t.isDark(n+1,o)&&t.isDark(n+2,o)&&t.isDark(n+3,o)&&t.isDark(n+4,o)&&!t.isDark(n+5,o)&&t.isDark(n+6,o)&&(e+=40);var l=0;for(o=0;o<r;o+=1)for(n=0;n<r;n+=1)t.isDark(n,o)&&(l+=1);return e+=10*(Math.abs(100*l/r/r-50)/5)},t}(),p=function(){for(var r=new Array(256),e=new Array(256),t=0;t<8;t+=1)r[t]=1<<t;for(t=8;t<256;t+=1)r[t]=r[t-4]^r[t-5]^r[t-6]^r[t-8];for(t=0;t<255;t+=1)e[r[t]]=t;var n={glog:function(t){if(t<1)throw"glog("+t+")";return e[t]},gexp:function(t){for(;t<0;)t+=255;for(;256<=t;)t-=255;return r[t]}};return n}();function m(n,o){if(void 0===n.length)throw n.length+"/"+o;var r=function(){for(var t=0;t<n.length&&0==n[t];)t+=1;for(var r=new Array(n.length-t+o),e=0;e<n.length-t;e+=1)r[e]=n[e+t];return r}(),i={getAt:function(t){return r[t]},getLength:function(){return r.length},multiply:function(t){for(var r=new Array(i.getLength()+t.getLength()-1),e=0;e<i.getLength();e+=1)for(var n=0;n<t.getLength();n+=1)r[e+n]^=p.gexp(p.glog(i.getAt(e))+p.glog(t.getAt(n)));return m(r,0)},mod:function(t){if(i.getLength()-t.getLength()<0)return i;for(var r=p.glog(i.getAt(0))-p.glog(t.getAt(0)),e=new Array(i.getLength()),n=0;n<i.getLength();n+=1)e[n]=i.getAt(n);for(n=0;n<t.getLength();n+=1)e[n]^=p.gexp(p.glog(t.getAt(n))+r);return m(e,0).mod(t)}};return i}function b(){var e=[],o={writeByte:function(t){e.push(255&t)},writeShort:function(t){o.writeByte(t),o.writeByte(t>>>8)},writeBytes:function(t,r,e){r=r||0,e=e||t.length;for(var n=0;n<e;n+=1)o.writeByte(t[n+r])},writeString:function(t){for(var r=0;r<t.length;r+=1)o.writeByte(t.charCodeAt(r))},toByteArray:function(){return e},toString:function(){var t="";t+="[";for(var r=0;r<e.length;r+=1)0<r&&(t+=","),t+=e[r];return t+="]"}};return o}var k,t,C=(k=[[1,26,19],[1,26,16],[1,26,13],[1,26,9],[1,44,34],[1,44,28],[1,44,22],[1,44,16],[1,70,55],[1,70,44],[2,35,17],[2,35,13],[1,100,80],[2,50,32],[2,50,24],[4,25,9],[1,134,108],[2,67,43],[2,33,15,2,34,16],[2,33,11,2,34,12],[2,86,68],[4,43,27],[4,43,19],[4,43,15],[2,98,78],[4,49,31],[2,32,14,4,33,15],[4,39,13,1,40,14],[2,121,97],[2,60,38,2,61,39],[4,40,18,2,41,19],[4,40,14,2,41,15],[2,146,116],[3,58,36,2,59,37],[4,36,16,4,37,17],[4,36,12,4,37,13],[2,86,68,2,87,69],[4,69,43,1,70,44],[6,43,19,2,44,20],[6,43,15,2,44,16],[4,101,81],[1,80,50,4,81,51],[4,50,22,4,51,23],[3,36,12,8,37,13],[2,116,92,2,117,93],[6,58,36,2,59,37],[4,46,20,6,47,21],[7,42,14,4,43,15],[4,133,107],[8,59,37,1,60,38],[8,44,20,4,45,21],[12,33,11,4,34,12],[3,145,115,1,146,116],[4,64,40,5,65,41],[11,36,16,5,37,17],[11,36,12,5,37,13],[5,109,87,1,110,88],[5,65,41,5,66,42],[5,54,24,7,55,25],[11,36,12,7,37,13],[5,122,98,1,123,99],[7,73,45,3,74,46],[15,43,19,2,44,20],[3,45,15,13,46,16],[1,135,107,5,136,108],[10,74,46,1,75,47],[1,50,22,15,51,23],[2,42,14,17,43,15],[5,150,120,1,151,121],[9,69,43,4,70,44],[17,50,22,1,51,23],[2,42,14,19,43,15],[3,141,113,4,142,114],[3,70,44,11,71,45],[17,47,21,4,48,22],[9,39,13,16,40,14],[3,135,107,5,136,108],[3,67,41,13,68,42],[15,54,24,5,55,25],[15,43,15,10,44,16],[4,144,116,4,145,117],[17,68,42],[17,50,22,6,51,23],[19,46,16,6,47,17],[2,139,111,7,140,112],[17,74,46],[7,54,24,16,55,25],[34,37,13],[4,151,121,5,152,122],[4,75,47,14,76,48],[11,54,24,14,55,25],[16,45,15,14,46,16],[6,147,117,4,148,118],[6,73,45,14,74,46],[11,54,24,16,55,25],[30,46,16,2,47,17],[8,132,106,4,133,107],[8,75,47,13,76,48],[7,54,24,22,55,25],[22,45,15,13,46,16],[10,142,114,2,143,115],[19,74,46,4,75,47],[28,50,22,6,51,23],[33,46,16,4,47,17],[8,152,122,4,153,123],[22,73,45,3,74,46],[8,53,23,26,54,24],[12,45,15,28,46,16],[3,147,117,10,148,118],[3,73,45,23,74,46],[4,54,24,31,55,25],[11,45,15,31,46,16],[7,146,116,7,147,117],[21,73,45,7,74,46],[1,53,23,37,54,24],[19,45,15,26,46,16],[5,145,115,10,146,116],[19,75,47,10,76,48],[15,54,24,25,55,25],[23,45,15,25,46,16],[13,145,115,3,146,116],[2,74,46,29,75,47],[42,54,24,1,55,25],[23,45,15,28,46,16],[17,145,115],[10,74,46,23,75,47],[10,54,24,35,55,25],[19,45,15,35,46,16],[17,145,115,1,146,116],[14,74,46,21,75,47],[29,54,24,19,55,25],[11,45,15,46,46,16],[13,145,115,6,146,116],[14,74,46,23,75,47],[44,54,24,7,55,25],[59,46,16,1,47,17],[12,151,121,7,152,122],[12,75,47,26,76,48],[39,54,24,14,55,25],[22,45,15,41,46,16],[6,151,121,14,152,122],[6,75,47,34,76,48],[46,54,24,10,55,25],[2,45,15,64,46,16],[17,152,122,4,153,123],[29,74,46,14,75,47],[49,54,24,10,55,25],[24,45,15,46,46,16],[4,152,122,18,153,123],[13,74,46,32,75,47],[48,54,24,14,55,25],[42,45,15,32,46,16],[20,147,117,4,148,118],[40,75,47,7,76,48],[43,54,24,22,55,25],[10,45,15,67,46,16],[19,148,118,6,149,119],[18,75,47,31,76,48],[34,54,24,34,55,25],[20,45,15,61,46,16]],(t={}).getRSBlocks=function(t,r){var e=function(t,r){switch(r){case w.L:return k[4*(t-1)+0];case w.M:return k[4*(t-1)+1];case w.Q:return k[4*(t-1)+2];case w.H:return k[4*(t-1)+3];default:return}}(t,r);if(void 0===e)throw"bad rs block @ typeNumber:"+t+"/errorCorrectionLevel:"+r;for(var n,o,i=e.length/3,a=[],u=0;u<i;u+=1)for(var f=e[3*u+0],c=e[3*u+1],l=e[3*u+2],g=0;g<f;g+=1)a.push((n=l,o=void 0,(o={}).totalCount=c,o.dataCount=n,o));return a},t),B=function(){var e=[],n=0,o={getBuffer:function(){return e},getAt:function(t){var r=Math.floor(t/8);return 1==(e[r]>>>7-t%8&1)},put:function(t,r){for(var e=0;e<r;e+=1)o.putBit(1==(t>>>r-e-1&1))},getLengthInBits:function(){return n},putBit:function(t){var r=Math.floor(n/8);e.length<=r&&e.push(0),t&&(e[r]|=128>>>n%8),n+=1}};return o},x=function(t){var r=a,n=t,e={getMode:function(){return r},getLength:function(t){return n.length},write:function(t){for(var r=n,e=0;e+2<r.length;)t.put(o(r.substring(e,e+3)),10),e+=3;e<r.length&&(r.length-e==1?t.put(o(r.substring(e,e+1)),4):r.length-e==2&&t.put(o(r.substring(e,e+2)),7))}},o=function(t){for(var r=0,e=0;e<t.length;e+=1)r=10*r+i(t.charAt(e));return r},i=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);throw"illegal char :"+t};return e},T=function(t){var r=u,n=t,e={getMode:function(){return r},getLength:function(t){return n.length},write:function(t){for(var r=n,e=0;e+1<r.length;)t.put(45*o(r.charAt(e))+o(r.charAt(e+1)),11),e+=2;e<r.length&&t.put(o(r.charAt(e)),6)}},o=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);if("A"<=t&&t<="Z")return t.charCodeAt(0)-"A".charCodeAt(0)+10;switch(t){case" ":return 36;case"$":return 37;case"%":return 38;case"*":return 39;case"+":return 40;case"-":return 41;case".":return 42;case"/":return 43;case":":return 44;default:throw"illegal char :"+t}};return e},M=function(t){var r=o,e=i.stringToBytes(t),n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=0;r<e.length;r+=1)t.put(e[r],8)}};return n},A=function(t){var r=f,n=i.stringToBytesFuncs.SJIS;if(!n)throw"sjis not supported.";!function(t,r){var e=n("友");if(2!=e.length||38726!=(e[0]<<8|e[1]))throw"sjis not supported."}();var o=n(t),e={getMode:function(){return r},getLength:function(t){return~~(o.length/2)},write:function(t){for(var r=o,e=0;e+1<r.length;){var n=(255&r[e])<<8|255&r[e+1];if(33088<=n&&n<=40956)n-=33088;else{if(!(57408<=n&&n<=60351))throw"illegal char at "+(e+1)+"/"+n;n-=49472}n=192*(n>>>8&255)+(255&n),t.put(n,13),e+=2}if(e<r.length)throw"illegal char at "+(e+1)}};return e},S=function(t){var e=t,n=0,o=0,i=0,r={read:function(){for(;i<8;){if(n>=e.length){if(0==i)return-1;throw"unexpected end of file./"+i}var t=e.charAt(n);if(n+=1,"="==t)return i=0,-1;t.match(/^\s$/)||(o=o<<6|a(t.charCodeAt(0)),i+=6)}var r=o>>>i-8&255;return i-=8,r}},a=function(t){if(65<=t&&t<=90)return t-65;if(97<=t&&t<=122)return t-97+26;if(48<=t&&t<=57)return t-48+52;if(43==t)return 62;if(47==t)return 63;throw"c:"+t};return r},L=function(t,r,e){for(var n=function(t,r){var n=t,o=r,g=new Array(t*r),e={setPixel:function(t,r,e){g[r*n+t]=e},write:function(t){t.writeString("GIF87a"),t.writeShort(n),t.writeShort(o),t.writeByte(128),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(255),t.writeByte(255),t.writeByte(255),t.writeString(","),t.writeShort(0),t.writeShort(0),t.writeShort(n),t.writeShort(o),t.writeByte(0);var r=i(2);t.writeByte(2);for(var e=0;255<r.length-e;)t.writeByte(255),t.writeBytes(r,e,255),e+=255;t.writeByte(r.length-e),t.writeBytes(r,e,r.length-e),t.writeByte(0),t.writeString(";")}},i=function(t){for(var r=1<<t,e=1+(1<<t),n=t+1,o=s(),i=0;i<r;i+=1)o.add(String.fromCharCode(i));o.add(String.fromCharCode(r)),o.add(String.fromCharCode(e));var a=b(),u=function(t){var e=t,n=0,o=0,r={write:function(t,r){if(t>>>r!=0)throw"length over";for(;8<=n+r;)e.writeByte(255&(t<<n|o)),r-=8-n,t>>>=8-n,n=o=0;o|=t<<n,n+=r},flush:function(){0<n&&e.writeByte(o)}};return r}(a);u.write(r,n);var f=0,c=String.fromCharCode(g[f]);for(f+=1;f<g.length;){var l=String.fromCharCode(g[f]);f+=1,o.contains(c+l)?c+=l:(u.write(o.indexOf(c),n),o.size()<4095&&(o.size()==1<<n&&(n+=1),o.add(c+l)),c=l)}return u.write(o.indexOf(c),n),u.write(e,n),u.flush(),a.toByteArray()},s=function(){var r={},e=0,n={add:function(t){if(n.contains(t))throw"dup key:"+t;r[t]=e,e+=1},size:function(){return e},indexOf:function(t){return r[t]},contains:function(t){return void 0!==r[t]}};return n};return e}(t,r),o=0;o<r;o+=1)for(var i=0;i<t;i+=1)n.setPixel(i,o,e(i,o));var a=b();n.write(a);for(var u=function(){function e(t){a+=String.fromCharCode(r(63&t))}var n=0,o=0,i=0,a="",t={},r=function(t){if(t<0);else{if(t<26)return 65+t;if(t<52)return t-26+97;if(t<62)return t-52+48;if(62==t)return 43;if(63==t)return 47}throw"n:"+t};return t.writeByte=function(t){for(n=n<<8|255&t,o+=8,i+=1;6<=o;)e(n>>>o-6),o-=6},t.flush=function(){if(0<o&&(e(n<<6-o),o=n=0),i%3!=0)for(var t=3-i%3,r=0;r<t;r+=1)a+="="},t.toString=function(){return a},t}(),f=a.toByteArray(),c=0;c<f.length;c+=1)u.writeByte(f[c]);return u.flush(),"data:image/gif;base64,"+u};return i}();a.stringToBytesFuncs["UTF-8"]=function(t){return function(t){for(var r=[],e=0;e<t.length;e++){var n=t.charCodeAt(e);n<128?r.push(n):n<2048?r.push(192|n>>6,128|63&n):n<55296||57344<=n?r.push(224|n>>12,128|n>>6&63,128|63&n):(e++,n=65536+((1023&n)<<10|1023&t.charCodeAt(e)),r.push(240|n>>18,128|n>>12&63,128|n>>6&63,128|63&n))}return r}(t)},o=[],void 0===(i="function"==typeof(n=function(){return a})?n.apply(r,o):n)||(t.exports=i)}])});