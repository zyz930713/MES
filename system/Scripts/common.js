// JavaScript Document
$(function() {
    $("input,select").bind("focus", function() {
        $(this).addClass("inputfocus");
    }).bind("blur", function() {
        $(this).removeClass("inputfocus");
    })
})