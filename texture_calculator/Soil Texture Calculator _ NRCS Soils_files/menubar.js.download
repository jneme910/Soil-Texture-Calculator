// Accessible Menubar widget 
// Source: Open AJAX Alliance
// http://oaa-accessibility.org/example/25/
// Modified by Terrill Thompson, http://staff.washington.edu/tft



var $s = {}
var isClick = false;
$s = jQuery.noConflict(); 



$s(document).ready(function() {
	var menu1 = new menubar('nav', false);
}); // end ready()

/////////////// begin menu widget definition /////////////////////
//
// Function menubar() is the constructor of a menu widget
// The widget will bind to the ul passed to it.
//
// @param(id string) id is the HTML id of the ul to bind to
//
// @param(vmenu boolean) vmenu is true if menu is vertical; false if horizontal
//
// @return N/A
//
function menubar(id, vmenu) {
	// define widget properties
	this.$sid = $s('#' + id);

	this.$srootItems = this.$sid.children('li'); // jQuerry array of all root-level menu items

	this.$sitems = this.$sid.find('.menu-item').not('.separator'); // jQuery array of menu items
	
	this.$sparents = this.$sid.find('.menu-parent'); // jQuery array of menu items

	this.$sallItems = this.$sparents.add(this.$sitems); // jQuery array of all menu items
	
	//console.log('All items:');
	//console.log(this.$sallItems);
	//console.log('Parents: ');
	//console.log(this.$sparents);
	
	
	this.$sactiveItem = null; // jQuery object of the menu item with focus

	this.vmenu = vmenu;
	this.bChildOpen = false; // true if child menu is open

	this.keys = {
		tab:    9,
		enter:  13,
		esc:    27,
		space:  32,
		left:   37,
		up:     38,
		right:  39,
		down:   40
	};

	// bind event handlers
	this.bindHandlers();

	// associate the menu with the textArea it controls
	// this.textarea = new textArea(this.$sid.attr('aria-controls'));
};

//
// Function bindHandlers() is a member function to bind event handlers for the widget.
//
// @return N/A
//
menubar.prototype.bindHandlers = function() {

	var thisObj = this;
	///////// bind mouse event handlers //////////
	// bind a handler for the menu items
	this.$sitems.mouseenter(function(e) {
		$s(this).addClass('menu-hover');
		return true;
	});

	// bind a mouseout handler for the menu items
	this.$sitems.mouseout(function(e) {
		$s(this).removeClass('menu-hover');
		return true;
	});

	// bind a mouseenter handler for the menu parents
	this.$sparents.mouseenter(function(e) {
		return thisObj.handleMouseEnter($s(this), e);
	});

	// bind a mouseleave handler
	this.$sparents.mouseleave(function(e) {
		return thisObj.handleMouseLeave($s(this), e);
	});

	// bind a click handler
	this.$sallItems.click(function(e) {
		return thisObj.handleClick($s(this), e);
	});

	//////////// bind key event handlers //////////////////
  
	// bind a keydown handler
	this.$sallItems.keydown(function(e) {
		return thisObj.handleKeyDown($s(this), e);
	});

	// bind a keypress handler
	this.$sallItems.keypress(function(e) {
		return thisObj.handleKeyPress($s(this), e);
	});

	// bind a focus handler
	this.$sallItems.focus(function(e) {
		return thisObj.handleFocus($s(this), e);
	});

	// bind a blur handler
	this.$sallItems.blur(function(e) {
		return thisObj.handleBlur($s(this), e);
	});
	//bind a document click within elements
	this.$sallItems.click(function(e){
		return thisObj.handleWithInDocumentClick($s(this),e);
	});
	// bind a document click handler
	$s(document).click(function(e) {
		return thisObj.handleDocumentClick(e);
	});

} // end bindHandlers()

//
// Function handleMouseEnter() is a member function to process mouseover
// events for the top menus.
//
// @param($sitem object) $sitem is the jquery object of the item firing the event
//
// @param(e object) e is the associated event object
//
// @return(boolean) Returns false;
//
menubar.prototype.handleMouseEnter = function($sitem, e) {
	
	// add hover style
	$sitem.addClass('menu-hover');
	// expand the first level submenu
	if ($sitem.attr('aria-haspopup') == 'true') {
		$sitem.children('ul').show().attr('aria-hidden', 'false');
		this.bChildOpen = true;
	}
	//e.stopPropagation();
	return true;

} // end handleMouseEnter()

//
// Function handleMouseOut() is a member function to process mouseout
// events for the top menus.
//
// @param($sitem object) $sitem is the jquery object of the item firing the event
//
// @param(e object) e is the associated event object
//
// @return(boolean) Returns false;
//
menubar.prototype.handleMouseOut = function($sitem, e) {
	
	// Remover hover styles
	$sitem.removeClass('menu-hover');

	//e.stopPropagation();
	return true;

} // end handleMouseOut()

//
// Function handleMouseLeave() is a member function to process mouseout
// events for the top menus.
//
// @param($smenu object) $smenu is the jquery object of the item firing the event
//
// @param(e object) e is the associated event object
//
// @return(boolean) Returns false;
//
menubar.prototype.handleMouseLeave = function($smenu, e) {

//console.log('mouseLeave');
	
 	
	var $sactive = $smenu.find('.menu-focus'); //???

	$sactive = $sactive.add($smenu.find('.menu-focus'));

	// Remove hover style
	$smenu.removeClass('menu-hover');

	// if any item in the child menu has focus, move focus to the root item
	if ($sactive.length > 0) {

		this.bChildOpen = false;

		// remove the focus style from the active item
		$sactive.removeClass('menu-focus');

		// store the active item
		this.$sactiveItem = $smenu;

		// cannot hide items with focus -- move focus to root item
		$smenu.focus();
	}

	// hide the child menu
	$smenu.children('ul').hide().attr('aria-hidden', 'true');
	
	//e.stopPropagation();
	return true;

} // end handleMouseLeave()

//
// Function handleClick() is a member function to process click events
// for the top menus.
//
// @param($sitem object) $sitem is the jquery object of the item firing the event
//
// @param(e object) e is the associated event object
//
// @return(boolean) Returns false;
//
menubar.prototype.handleClick = function($sitem, e) {
	
	var $sparentUL = $sitem.parent();
		
		if($sitem.hasClass('sub-menu-parent') || ($sitem.hasClass('center_nav'))){
			//This check if for the sub menu list opening up. Making sure that nothing else changes.
			//do Nothing

		}else{
		
		if ($sparentUL.is('.root-level')) {
			// open the child menu if it is closed
			$sitem.children('ul').first().show().attr('aria-hidden', 'false');
			this.bChildOpen = true;
		}
		else {
			// remove hover and focus styling
			this.$sallItems.removeClass('menu-hover menu-focus');

			// close the menu
			this.$sid.find('ul').not('.root-level').hide().attr('aria-hidden','true');
		}

		// if menu item triggers some behavior other than going to a link, 
		// would stop propagation and return false 
		// e.stopPropagation();
		// return false;
		return true;
	}
} // end handleClick()

//
// Function handleFocus() is a member function to process focus events
// for the menu.
//
// @param($sitem object) $sitem is the jquery object of the item firing the event
//
// @param(e object) e is the associated event object
//
// @return(boolean) Returns true;
//
menubar.prototype.handleFocus = function($sitem, e) {
	
	// if activeItem is null, we are getting focus from outside the menu. Store
	// the item that triggered the event
	if (this.$sactiveItem == null) {
		this.$sactiveItem = $sitem;
	}
	else if ($sitem[0] != this.$sactiveItem[0]) {
		return true;
	}
  
	// get the set of jquery objects for all the parent items of the active item
	var $sparentItems = this.$sactiveItem.parentsUntil('div').filter('li');

	// remove focus styling from all other menu items
	this.$sallItems.removeClass('menu-focus');

	// add styling to the active item 
	this.$sactiveItem.addClass('menu-focus');
/*	
	if (this.$sactiveItem.hasClass('menu-parent')) { 
		// for parent items, add .menu-focus directly to the list item
		this.$sactiveItem.addClass('menu-focus');
	}
	else { 
		// for sub-menu items, add .menu-focus to the anchor
		this.$sactiveItem.find('a').addClass('menu-focus');
	}
*/
	// add styling to all parent items.
	$sparentItems.addClass('menu-focus');

	if (this.vmenu == true) {
		// if the bChildOpen is true, open the active item's child menu (if applicable)
		if (this.bChildOpen == true) {

			var $sitemUL = $sitem.parent();

			// if the itemUL is a root-level menu and item is a parent item,
			// show the child menu.
			if ($sitemUL.is('.root-level') && ($sitem.attr('aria-haspopup') == 'true')) {
				$sitem.children('ul').show().attr('aria-hidden', 'false');
			}
		}
		else {
			this.vmenu = false;
		}
	}

	return true;

} // end handleFocus()

//
// Function handleBlur() is a member function to process blur events
// for the menu.
//
// @param($sitem object) $sitem is the jquery object of the item firing the event
//
// @param(e object) e is the associated event object
//
// @return(boolean) Returns true;
//
menubar.prototype.handleBlur = function($sitem, e) {
	
	// $sitem.find('a').removeClass('menu-focus');
	$sitem.removeClass('menu-focus');

	return true;

} // end handleBlur()

//
// Function handleKeyDown() is a member function to process keydown events
// for the menus.
//
// @param($sitem object) $sitem is the jquery object of the item firing the event
//
// @param(e object) e is the associated event object
//
// @return(boolean) Returns false if consuming; true if propagating
//
menubar.prototype.handleKeyDown = function($sitem, e) {
	
	if (e.altKey || e.ctrlKey) {
		// Modifier key pressed: Do not process
		return true;
	}

	switch(e.keyCode) {
		case this.keys.tab: {

			// hide all menu items and update their aria attributes
			this.$sid.find('ul').hide().attr('aria-hidden', 'true');
		    
			
			// remove focus styling from all menu items
			this.$sallItems.removeClass('menu-focus');

			this.$sactiveItem = null;
			this.bChildOpen == false;
				
				//making sure tabbing after menu bar goes to the additional links in the top menu.
			if($sitem.attr('class') == 'menu-parent' && !e.shiftKey){
				
				if($sitem.find('ul').attr('id') == 'specialId'){
					
					//e.stopPropagation();
					//$s('a[tabindex=14]').focus();
					$s('.addLink_First').focus();
					return false;
					//$s('.addLink_First').focus();
				}
			
			}
			
			
			break;
		}
		case this.keys.esc: {

			var $sitemUL = $sitem.parent();

			if ($sitemUL.is('.root-level')) {
				// hide the child menu and update the aria attributes
				$sitem.children('ul').first().hide().attr('aria-hidden', 'true');
			}
			else {

				// move up one level
				this.$sactiveItem = $sitemUL.parent();

				// reset the childOpen flag
				this.bChildOpen = false;

				// set focus on the new item
				this.$sactiveItem.focus();

				// hide the active menu and update the aria attributes
				$sitemUL.hide().attr('aria-hidden', 'true');
			}

			e.stopPropagation();
			return false;
		}
		case this.keys.space:{
			
			//For opening Sub Menu Items under the Topics Menu.
			if($sitem.hasClass('sub-menu-parent')){
				
				$sitem.find('ul').show().attr('aria-hidden','false');
				
				var $snewItemUL = $sitem.children('ul').first();
				var $snewItem = $snewItemUL.children('li').first();
				
				this.$sactiveItem = $snewItem;
				
				this.$sactiveItem.focus();

				e.stopPropagation();				
				return false;
			}
			
			if (!($sitem.hasClass('menu-parent'))) { 
				// user pressed enter or space on a dropdown menu item, 
				// not an item on the menu bar 		
				// get the target href and go there		
				//window.location = $sitem.find('a').attr('href');	
				return false;
			}
			
			
			
			var $sparentUL = $sitem.parent();

			if ($sparentUL.is('.root-level')) {
				// open the child menu if it is closed
				$sitem.children('ul').first().show().attr('aria-hidden', 'false');
				$sitem.children('ul').first().children('li').first().focus();
				$sitem.children('ul').first().children('li').first().addClass('menu-focus');
				this.bChildOpen = true;
			}
			else {
				// remove hover styling
				this.$sallItems.removeClass('menu-hover');
				this.$sallItems.removeClass('menu-focus');

				// close the menu
				this.$sid.find('ul').not('.root-level').hide().attr('aria-hidden','true');

				// clear the active item
				this.$sactiveItem = null;

			}
			
			e.stopPropagation();
			return false;
			
		}
		case this.keys.enter: {
				
			if($sitem.hasClass('sub-menu-parent')){
				window.location = $sitem.find('a').not('.expand_link').attr('href');
				return false;
			}			
				
			//if (!($sitem.hasClass('menu-parent'))) { 
				// user pressed enter or space on a dropdown menu item, 
				// not an item on the menu bar 		
				// get the target href and go there		
				window.location = $sitem.find('a').attr('href');	
				return false;
			//}

			var $sparentUL = $sitem.parent();

			if ($sparentUL.is('.root-level')) {
				// open the child menu if it is closed
				$sitem.children('ul').first().show().attr('aria-hidden', 'false');
				this.bChildOpen = true;
			}
			else {
				// remove hover styling
				this.$sallItems.removeClass('menu-hover');
				this.$sallItems.removeClass('menu-focus');

				// close the menu
				this.$sid.find('ul').not('.root-level').hide().attr('aria-hidden','true');

				// clear the active item
				this.$sactiveItem = null;

			}
			e.stopPropagation();
			return false;
		}

		case this.keys.left: {

			if (this.vmenu == true && $sitemUL.is('.root-level')) {
				// If this is a vertical menu and the root-level is active, move
				// to the previous item in the menu
				this.$sactiveItem = this.moveUp($sitem);
			}
			else {
				this.$sactiveItem = this.moveToPrevious($sitem);
			}

			this.$sactiveItem.focus();

			e.stopPropagation();
			return false;
		}
		case this.keys.right: {

			if (this.vmenu == true && $sitemUL.is('.root-level')) {
				// If this is a vertical menu and the root-level is active, move
				// to the next item in the menu
				this.$sactiveItem = this.moveDown($sitem);
			}
			else {
				this.$sactiveItem = this.moveToNext($sitem);
			}

			this.$sactiveItem.focus();

			e.stopPropagation();
			return false;
		}
		case this.keys.up: {

			if (this.vmenu == true && $sitemUL.is('.root-level')) {
				// If this is a vertical menu and the root-level is active, move
				// to the previous root-level menu
				this.$sactiveItem = this.moveToPrevious($sitem);
			}
			else {
				this.$sactiveItem = this.moveUp($sitem);
			}

			this.$sactiveItem.focus();

			e.stopPropagation();
			return false;
		}
		case this.keys.down: {

			if (this.vmenu == true && $sitemUL.is('.root-level')) {
				// If this is a vertical menu and the root-level is active, move
				// to the next root-level menu
				this.$sactiveItem = this.moveToNext($sitem);
			}
			else {
				this.$sactiveItem = this.moveDown($sitem);
			}

			this.$sactiveItem.focus();

			e.stopPropagation();
			return false;
		}
	} // end switch
	return true;

} // end handleKeyDown()

//
// Function moveToNext() is a member function to move to the next menu level.
// This will be either the next root-level menu or the child of a menu parent. If
// at the root level and the active item is the last in the menu, this function will loop
// to the first menu item.
//
// If the menu is a horizontal menu, the first child element of the newly selected menu will
// be selected
//
// @param($sitem object) $sitem is the active menu item
//
// @return (object) Returns the item to move to. Returns $sitem is no move is possible
//
menubar.prototype.moveToNext = function($sitem) {

	var $sitemUL = $sitem.parent(); // $sitem's containing menu
	var $smenuItems = $sitemUL.children('li'); // the items in the currently active menu
	var menuNum = $smenuItems.length; // the number of items in the active menu
	var menuIndex = $smenuItems.index($sitem); // the items index in its menu
	var $snewItem = null;
	var $snewItemUL = null;

	if ($sitemUL.is('.root-level')) {
		// this is the root level move to next sibling. This will require closing
		// the current child menu and opening the new one.

		if (menuIndex < menuNum-1) { // not the last root menu
			
			$snewItem = $sitem.next();
			//If condition added to skip the transparent li item.
			if($snewItem.hasClass('transparent_nav')){
				$snewItem = $snewItem.next();
			}else if($snewItem.hasClass('right_corner_nav')){
				$snewItem = $smenuItems.first().next();
			}
		}
		else { // wrap to first item
			$snewItem = $smenuItems.first();
		}

		// close the current child menu (if applicable)
		if ($sitem.attr('aria-haspopup') == 'true') {

			var $schildMenu = $sitem.children('ul').first();

			if ($schildMenu.attr('aria-hidden') == 'false') {
				// hide the child and update aria attributes accordingly
				$schildMenu.hide().attr('aria-hidden', 'true');
				this.bChildOpen = true;
			}
		}

		// remove the focus styling from the current menu
		$sitem.removeClass('menu-focus');

		// open the new child menu (if applicable)
		if (($snewItem.attr('aria-haspopup') == 'true') && (this.bChildOpen == true)) {

			var $schildMenu = $snewItem.children('ul').first();

			// open the child and update aria attributes accordingly
			$schildMenu.show().attr('aria-hidden', 'false');

			if (!this.vmenu) {
				// select the first item in the child menu
				$snewItem = $schildMenu.children('li').first();
			}

		}
	}
	else {
		// this is not the root level. If there is a child menu to be moved into, do that;
		// otherwise, move to the next root-level menu if there is one
		if ($sitem.attr('aria-haspopup') == 'true') {
        
			var $schildMenu = $sitem.children('ul').first();

			$snewItem = $schildMenu.children('li').first();

			// show the child menu and update its aria attributes
			$schildMenu.show().attr('aria-hidden', 'false');
			this.bChildOpen = true;
		}
		else {
			// at deepest level, move to the next root-level menu

			if (this.vmenu == true) {
				// do nothing
				return $sitem;
			}

			var $sparentMenus = null;
			var $srootItem = null;

			// get list of all parent menus for item, up to the root level
			$sparentMenus = $sitem.parentsUntil('div').filter('ul').not('.root-level');

			// hide the current menu and update its aria attributes accordingly
			$sparentMenus.hide().attr('aria-hidden', 'true');

			// remove the focus styling from the active menu
			$sparentMenus.find('li').removeClass('menu-focus');
			$sparentMenus.last().parent().removeClass('menu-focus');

			$srootItem = $sparentMenus.last().parent(); // the containing root for the menu

			menuIndex = this.$srootItems.index($srootItem);

			// if this is not the last root menu item, move to the next one
			if (menuIndex < this.$srootItems.length-1) {
				$snewItem = $srootItem.next();
				//If condition added to skip the transparent li item.
				if($snewItem.hasClass('transparent_nav')){
					$snewItem = $snewItem.next();
				}else if($snewItem.hasClass('right_corner_nav')){
					$snewItem = this.$srootItems.first().next();
				}
			}
			else { // loop
				$snewItem = this.$srootItems.first();
			}

			if ($snewItem.attr('aria-haspopup') == 'true') {
				var $schildMenu = $snewItem.children('ul').first();

				$snewItem = $schildMenu.children('li').first();

				// show the child menu and update it's aria attributes
				$schildMenu.show().attr('aria-hidden', 'false');
				this.bChildOpen = true;
			}
		}
	}

	return $snewItem;
}

//
// Function moveToPrevious() is a member function to move to the previous menu level.
// This will be either the previous root-level menu or the child of a menu parent. If
// at the root level and the active item is the first in the menu, this function will loop
// to the last menu item.
//
// If the menu is a horizontal menu, the first child element of the newly selected menu will
// be selected
//
// @param($sitem object) $sitem is the active menu item
//
// @return (object) Returns the item to move to. Returns $sitem is no move is possible
//
menubar.prototype.moveToPrevious = function($sitem) {

	var $sitemUL = $sitem.parent(); // $sitem's containing menu
	var $smenuItems = $sitemUL.children('li'); // the items in the currently active menu
	var menuNum = $smenuItems.length; // the number of items in the active menu
	var menuIndex = $smenuItems.index($sitem); // the items index in its menu
	var $snewItem = null;
	var $snewItemUL = null;

	if ($sitemUL.is('.root-level')) {
		// this is the root level move to previous sibling. This will require closing
		// the current child menu and opening the new one.

		if (menuIndex > 0) { // not the first root menu
			$snewItem = $sitem.prev();
			//If condition added to skip the transparent li item and left Corner li
			if($snewItem.hasClass('transparent_nav')){
				$snewItem = $snewItem.prev();
			}else if($snewItem.hasClass('left_corner_nav')){
				$snewItem = $smenuItems.last().prev();
			}
		}
		else { // wrap to last item
			$snewItem = $smenuItems.last();
		}
		
		// close the current child menu (if applicable)
		if ($sitem.attr('aria-haspopup') == 'true') {

			var $schildMenu = $sitem.children('ul').first();

			if ($schildMenu.attr('aria-hidden') == 'false') {
				// hide the child and update aria attributes accordingly
				$schildMenu.hide().attr('aria-hidden', 'true');
				this.bChildOpen = true;
			}
		}

		// remove the focus styling from the current menu
		$sitem.removeClass('menu-focus');

		// open the new child menu (if applicable)
		if (($snewItem.attr('aria-haspopup') == 'true') && this.bChildOpen == true) {

			var $schildMenu = $snewItem.children('ul').first();

			// open the child and update aria attributes accordingly
			$schildMenu.show().attr('aria-hidden', 'false');

			if (!this.vmenu) {
				// select the first item in the child menu
				$snewItem = $schildMenu.children('li').first();
			}
		}
	}
	else {
		// this is not the root level. If there is a parent menu that is not the
		// root menu, move up one level; otherwise, move to first item of the previous
		// root menu.

		var $sparentLI = $sitemUL.parent();
		var $sparentUL = $sparentLI.parent();

		var $sparentMenus = null;
		var $srootItem = null;

		// if this is a vertical menu or is not the first child menu
		// of the root-level menu, move up one level.
		if (this.vmenu == true || !$sparentUL.is('.root-level')) {

			$snewItem = $sitemUL.parent();

			// hide the active menu and update aria-hidden
			$sitemUL.hide().attr('aria-hidden', 'true');

			// remove the focus highlight from the $sitem
			$sitem.removeClass('menu-focus');

			if (this.vmenu == true) {
				// set a flag so the focus handler does't reopen the menu
				this.bChildOpen = false;
			}
		}
		else { // move to previous root-level menu

			// hide the current menu and update the aria attributes accordingly
			$sitemUL.hide().attr('aria-hidden', 'true');

			// remove the focus styling from the active menu
			$sitem.removeClass('menu-focus');
			$sparentLI.removeClass('menu-focus');

			menuIndex = this.$srootItems.index($sparentLI);

			if (menuIndex > 0) {
				// move to the previous root-level menu
				$snewItem = $sparentLI.prev();
				//If condition added to skip the transparent li item.
				if($snewItem.hasClass('transparent_nav')){
					$snewItem = $snewItem.prev();
				
				}else if($snewItem.hasClass('left_corner_nav')){
						$snewItem = this.$srootItems.last().prev();
				}
			}
			else { // loop to last root-level menu
				$snewItem = this.$srootItems.last();
				
			}

			// add the focus styling to the new menu
			$snewItem.addClass('menu-focus');

			if ($snewItem.attr('aria-haspopup') == 'true') {
				var $schildMenu = $snewItem.children('ul').first();

				// show the child menu and update it's aria attributes
				$schildMenu.show().attr('aria-hidden', 'false');
				this.bChildOpen = true;

				$snewItem = $schildMenu.children('li').first();
			}
		}
	}
	return $snewItem;
}

//
// Function moveDown() is a member function to select the next item in a menu.
// If the active item is the last in the menu, this function will loop to the
// first menu item.
//
// @param($sitem object) $sitem is the active menu item
//
// @param(startChr char) [optional] startChr is the character to attempt to match against the beginning of the
// menu item titles. If found, focus moves to the next menu item beginning with that character.
//
// @return (object) Returns the item to move to. Returns $sitem is no move is possible
//
menubar.prototype.moveDown = function($sitem, startChr) {

	var $sitemUL = $sitem.parent(); // $sitem's containing menu
	var $smenuItems = $sitemUL.children('li').not('.separator'); // the items in the currently active menu
	var menuNum = $smenuItems.length; // the number of items in the active menu
	var menuIndex = $smenuItems.index($sitem); // the items index in its menu
	var $snewItem = null;
	var $snewItemUL = null;

	if ($sitemUL.is('.root-level')) { // this is the root level menu

		if ($sitem.attr('aria-haspopup') != 'true') {
			// No child menu to move to
			return $sitem;
		}

		// Move to the first item in the child menu
		$snewItemUL = $sitem.children('ul').first();
		$snewItem = $snewItemUL.children('li').first();

		// make sure the child menu is visible
		$snewItemUL.show().attr('aria-hidden', 'false');
		this.bChildOpen = true;

		return $snewItem;
	}

	// if $sitem is not the last item in its menu, move to the next item. If startChr is specified, move
	// to the next item with a title that begins with that character.
	//
	if (startChr) {
	
		var bMatch = false;
		var curNdx = menuIndex+1;

		// check if the active item was the last one on the list
		if (curNdx == menuNum) {
			curNdx = 0;
		}

		// Iterate through the menu items (starting from the current item and wrapping) until a match is found
		// or the loop returns to the current menu item
		while (curNdx != menuIndex)  {

			// Use the first of the two following lines if menu does not contain anchor tags. 
			// Otherwise use the second 
			// var titleChr = $smenuItems.eq(curNdx).html().charAt(0);
			var titleChr = $smenuItems.eq(curNdx).find('a').html().charAt(0);
			
			if (titleChr.toLowerCase() == startChr) {
				bMatch = true;
				break;
			}

			curNdx = curNdx+1;

			if (curNdx == menuNum) {
				// reached the end of the list, start again at the beginning
				curNdx = 0;
			}
		}

		if (bMatch == true) {
			$snewItem = $smenuItems.eq(curNdx);

			// remove the focus styling from the current item
			$sitem.removeClass('menu-focus');

			return $snewItem
		}
		else {
			return $sitem;
		}
	}
	else {
		if (menuIndex < menuNum-1) {
			$snewItem = $smenuItems.eq(menuIndex+1);
		}
		else {
			$snewItem = $smenuItems.first();
		}
	}

	// remove the focus styling from the current item
	$sitem.removeClass('menu-focus');

	return $snewItem;
}

//
// Function moveUp() is a member function to select the previous item in a menu.
// If the active item is the first in the menu, this function will loop to the
// last menu item.
//
// @param($sitem object) $sitem is the active menu item
//
// @return (object) Returns the item to move to. Returns $sitem is no move is possible
//
menubar.prototype.moveUp = function($sitem) {

	var $sitemUL = $sitem.parent(); // $sitem's containing menu
	var $smenuItems = $sitemUL.children('li').not('.separator'); // the items in the currently active menu
	var menuNum = $smenuItems.length; // the number of items in the active menu
	var menuIndex = $smenuItems.index($sitem); // the items index in its menu
	var $snewItem = null;
	var $snewItemUL = null;

	if ($sitemUL.is('.root-level')) { // this is the root level menu

		// nothing to do
		return $sitem;
	}

	// if $sitem is not the first item in its menu, move to the previous item
	if (menuIndex > 0) {

		$snewItem = $smenuItems.eq(menuIndex-1);
	}
	else {
		// loop to top of menu
		$snewItem = $smenuItems.last();
	}

	// remove the focus styling from the current item
	$sitem.removeClass('menu-focus');

	return $snewItem;
}

//
// Function handleKeyPress() is a member function to process keydown events
// for the menus.
//
// The Opera browser performs some window commands from the keypress event,
// not keydown like Firefox, Safari, and IE. This event handler consumes
// keypresses for relevant keys so that Opera behaves when the user is
// manipulating the menu with the keyboard.
//
// @param($sitem object) $smenu is the jquery object of the item firing the event
//
// @param(e object) e is the associated event object
//
// @return(boolean) Returns false if consuming; true if propagating
//
menubar.prototype.handleKeyPress = function($sitem, e) {


	if (e.altKey || e.ctrlKey || e.shiftKey) {
		// Modifier key pressed: Do not process
		return true;
	}

	switch(e.keyCode) {
		case this.keys.tab: {
			return true;
		}
		case this.keys.esc:
		case this.keys.enter: 
		case this.keys.space:
		case this.keys.up:
		case this.keys.down:
		case this.keys.left:
		case this.keys.right: {

			e.stopPropagation();
			return false;
		}
		default : {
			var chr = String.fromCharCode(e.which);

			this.$sactiveItem = this.moveDown($sitem, chr);
			this.$sactiveItem.focus();

			e.stopPropagation();
			return false;
		}
	} // end switch
	return true;

} // end handleKeyPress()

menubar.prototype.handleWithInDocumentClick = function($sitem,e){
	
	if($sitem.hasClass('sub-menu-parent')){
		e.stopPropagation();
	}
}
//
// Function handleDocumentClick() is a member function to process click events on the document. Needed
// to close an open menu if a user clicks outside the menu
//
// @param(e object) e is the associated event object
//
// @return(boolean) Returns true;
//
menubar.prototype.handleDocumentClick = function(e) {
	
	// get a list of all child menus
	var $schildMenus = this.$sid.find('ul').not('.root-level');

	// hide the child menus
	$schildMenus.hide().attr('aria-hidden', 'true');

	this.$sallItems.removeClass('menu-focus');

	this.$sactiveItem = null;

	// allow the event to propagate
	return true;
	
} // end handleDocumentClick()


/////////////// end menu widget definition /////////////////////	