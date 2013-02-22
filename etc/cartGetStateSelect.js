function m() {
  var country,countryId,stateId,name,cs,id,element,abbreviation,isFound;
  m='';
  element=cp.doc.Var('element');
  country=cp.doc.Var('country');
  country=country.replace('+',' '); /* old bug */
  countryId=cp.content.getRecordId('countries',country)
  state=cp.doc.Var('state').toLowerCase();;
  isFound=false;
  cs=cp.csNew();
  cs.open("states","countryid="+cp.db.encodeSqlNumber(countryId),"name");
  if(!cs.ok()){
    m+='<input type="text" name="'+element+'" value="'+state+'">';
    //m+='(state not required)';
  }else{
    while (cs.ok()) {
      id=cs.getInteger('id');
      name=cs.getText('name');
      abbreviation=cs.getText('abbreviation');
      if((name.toLowerCase()==state)||(abbreviation.toLowerCase()==state)){
        m+='\n\t\t<option selected>'+name+'</option>';
		isFound=true;
      }else{
        m+='\n\t\t<option>'+name+'</option>';
      }
      cs.goNext();
    }
	if(!isFound)
	{
        m+='\n\t\t<option value=\"\" selected>Select State</option>';
	}
    m='\n\t<select id="'+element+'" name="'+element+'">'+m+'\n\t</select>'
  }
  cs.close();
  return m;
}