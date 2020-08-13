// JavaScript Document
function showtable(tablename) {
  var color1 = "#ffffff";
  var color2 = "#ffffff";
  var bgColor = "#f8f8f8";
  var trs = document.getElementById(tablename).getElementsByTagName("tr");
  for (var i = 0; i < trs.length; i++) {
    if (i % 2 == 0) {
      trs[i].style.backgroundColor = color1;
      trs[i].onmouseover = function() {
        this.style.backgroundColor = bgColor;
      }
      trs[i].onmouseout = function() {
        this.style.backgroundColor = color1;
      }
    } else {
      trs[i].style.backgroundColor = color2;
      trs[i].onmouseover = function() {
        this.style.backgroundColor = bgColor;
      }
      trs[i].onmouseout = function() {
        this.style.backgroundColor = color2;
      }
    }
  }
}

function showtable2(tablename) {
  var color1 = "#ffffff";
  var color2 = "#ffffff";
  var bgColor = "#f8f8f8";
  var trs = document.getElementById(tablename).getElementsByTagName("tr");
  for (var i = 0; i < trs.length; i++) {
    if (i % 2 == 0) {
      trs[i].style.backgroundColor = color1;
      trs[i].onmouseover = function() {
        this.style.backgroundColor = bgColor;
      }
      trs[i].onmouseout = function() {
        this.style.backgroundColor = color1;
      }
    } else {
      trs[i].style.backgroundColor = color2;
      trs[i].onmouseover = function() {
        this.style.backgroundColor = bgColor;
      }
      trs[i].onmouseout = function() {
        this.style.backgroundColor = color2;
      }
    }
  }
}