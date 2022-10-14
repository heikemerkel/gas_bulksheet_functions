////////////////////////////////////////////////////////////////////////////
// AlaSQLGS 
// Script ID: 1XWR3NzQW6fINaIaROhzsxXqRREfKXAdbKoATNbpygoune43oCmez1N8U --> load as library
// (https://freesoft.dev/program/89807704)
////////////////////////////////////////////////////////////////////////////
function alaSQL(){
  const alasql = AlaSQLGS.load();
  const query = "SELECT MATRIX [0] FROM ? WHERE [0] LIKE ('" +name+"%')";  //first column is [0]
  console.log("query: ", query);
  var res = alasql(AlaSQLGS.transformQueryColsNotation(query), [values]);
  var res1 = alasql('SELECT [0] FROM ?',[values]);
}