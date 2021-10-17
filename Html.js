function getWithIndex(html, tag, index) {  
  var positionStart = html.indexOf(tag);
  if (positionStart !== -1) {
    var splittedHtmlArray = html.split(tag);
    
    var stringToFind = splittedHtmlArray[index];
    
    if(stringToFind == undefined)
      return false;
      
    var positionEnd = stringToFind.indexOf('</');
    
    var data = stringToFind.substring(0, positionEnd);
    return data;
  }
}

function getWithoutIndex(html, tag) {
  var positionStart = html.indexOf(tag);
  if (positionStart !== -1) {
    var positionEnd = html.indexOf('</',positionStart);
    
    var data = html.substring(positionStart + tag.length, positionEnd);
    return data;
  }
}

function getFloat(html, tag, index) {
  var data = getWithIndex(html, tag, index);
  
  if (data) {
    return parseFloat(data.replace(',','.'));
  }
}

function getDateBySeparator(html, tag, index, separator) {
  var data = getWithIndex(html, tag, index);
  
  if (data) {
    var splittedDate = data.split(separator);
    
    if(splittedDate.length == 1)
      return false;
    
    var dd = parseFloat(splittedDate[0]);
    var mm = parseFloat(splittedDate[1]) - 1;
    
    var year = splittedDate[2];
    if (year.length == 2){
      year = '20' + year;
    }
    var yyyy = parseFloat(year);
    
    var dateObject = new Date(yyyy, mm, dd);
    
    return dateObject;
  }
}