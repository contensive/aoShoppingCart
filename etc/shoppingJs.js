//
//----------
//
function toggleHide(isChecked,divId) {
	var e=document.getElementById(divId);
	if(isChecked){
		e.style.display='none';
	}else{
		e.style.display='block';
	}
}
//
//----------
//
function cartResetShippingSelect(){
      var el = document.getElementById('shipContainer');
      el.innerHTML = '<select onClick="cartSetImg();cartDrawShippingSelect();"><option>Select One</option></select>';
}
//
//----------
//
function cartSetImg(){
      var el = document.getElementById('shipContainer');
      el.innerHTML = '<img src="/ShoppingCart/loadSpinner.gif">';
}
//
//----------
//
function cartDrawShippingSelect(){
	var elNM = document.getElementById('shipMethod');
	if( elNM ){
		var elCB = document.getElementById('ShipToContactAddress');
		var elWT = document.getElementById('shipWeight');
		var elCH = document.getElementById('shipCharge');
		var elZP;
		var elCT;
		var reqStr = 'shipMethod=' + elNM.value + '&shipWeight=' + elWT.value + '&shipCharge=' + elCH.value;
		var zip,country;

		if (elCB.checked) {
			elZP = document.getElementById('orderContactZip');
			zip = elZP.value;
			elCT = document.getElementById("orderContactCountryValue");
			if (elCT ) {
				country = elCT.value;
			} else {
				elCT = document.getElementById("orderContactCountry");	
				country = elCT.options[elCT.selectedIndex].value;
			}
		} else {
			elZP = document.getElementById('orderShipZip');
			elCT = document.getElementById('orderShipCountry');
			zip = elZP.value;
			country = elCT.options[elCT.selectedIndex].value;
		}
		cartSetImg();
		reqStr += '&shipZip=' + zip + '&shipCountry=' + country;
		cj.ajax.addonCallback('checkoutShipResponder',reqStr,cartDrawShippingSelectCallback,'shipContainer');
	}
}
//
//----------
//
function cartDrawShippingSelectCallback( response, destinationId ) {
    //alert( 'response='+response );
    //alert( 'destinationId='+destinationId );
    jQuery( '#'+destinationId ).html( response );
    jQuery( '#'+destinationId+' select' ).attr('size','5');
}

//
//----------
//
function cartInsertRow() {
    var RLTable = document.getElementById('scCheckoutTable');
    var CountElement = document.getElementById('CartRows');
    var tCols,tRows,NewRowNumber;
    var NewRow,e,NewCell2,NewCell3,NewCell4;
    if ( RLTable ) {
      NewRowNumber = parseInt( CountElement.value );
      CountElement.value = NewRowNumber+1;
      tRows = RLTable.getElementsByTagName("TR");
      NewRow = RLTable.insertRow(-1);
      e = NewRow.insertCell(-1);
      e.align = 'center';
      e.style.color='black';
      e.style.backgroundColor='white';
      e.style.borderBottom='1px solid #a0a0a0';
      e.style.borderRight='1px solid #a0a0a0';
      e.style.paddingRight='5px';
      e.style.paddingLeft='5px';
      e.innerHTML= '<INPUT TYPE=TEXT NAME=Q'+NewRowNumber+' VALUE=1 SIZE=2 MAXLENGTH=5>';
      e = NewRow.insertCell(-1);
      e.style.color='black';
      e.style.backgroundColor='white';
      e.style.borderBottom='1px solid #a0a0a0';
      e.style.borderRight='1px solid #a0a0a0';
      e.style.paddingRight='5px';
      e.style.paddingLeft='5px';
      e.innerHTML=ItemSelect.replace('row0id','ID'+NewRowNumber);
      NewCell3 = NewRow.insertCell(-1);
      NewCell3.style.color='black';
      NewCell3.style.backgroundColor='white';
      NewCell3.style.borderBottom='1px solid #a0a0a0';
      NewCell3.style.borderRight='1px solid #a0a0a0';
      NewCell3.style.paddingRight='5px';
      NewCell3.style.paddingLeft='5px';
      NewCell3.innerHTML='&nbsp;';
      NewCell4 = NewRow.insertCell(-1);
      NewCell4.style.color='black';
      NewCell4.style.backgroundColor='white';
      NewCell4.style.borderBottom='1px solid #a0a0a0';
      NewCell4.style.paddingRight='5px';
      NewCell4.style.paddingLeft='5px';
      NewCell4.innerHTML='&nbsp;';
    }
}
