/* Info */
var w = window.innerWidth
    || document.documentElement.clientWidth
    || document.body.clientWidth;

var h = window.innerHeight
    || document.documentElement.clientHeight
    || document.body.clientHeight;
/*Proccess*/
var Percent = (h - 135) / 100;
/*Define*/
var HeightTopBar = Percent * 20;
var HeightMain = Percent * 80;
var HeightFooter = Percent * 20;
/*Set*/
document.getElementById("Top").style.height = HeightTopBar + "px";
document.getElementById("Main").style.height = HeightMain + "px";
document.getElementById("Footer").style.height = HeightFooter + "px";
/*Event*/
window.addEventListener("resize", changeResize);
function changeResize() {
    document.getElementById("Top").style.height = HeightTopBar + "px";
    document.getElementById("Main").style.height = HeightMain + "px";
    document.getElementById("Footer").style.height = HeightFooter + "px";

    document.getElementById("Top").innerHTML = HeightTopBar;
    document.getElementById("Main").innerHTML = HeightMain;
    document.getElementById("Footer").innerHTML = HeightFooter;
}