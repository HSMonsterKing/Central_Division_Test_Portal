function switchtext(e) {
    if(document.getElementById(e.id.replace("Button1", "Label3")).innerHTML == ""){
        tb6 = document.getElementById(e.id.replace("Button1", "TextBox6")).value;
        tb7 = document.getElementById(e.id.replace("Button1", "TextBox7")).value;
        document.getElementById(e.id.replace("Button1", "TextBox6")).value = document.getElementById(e.id.replace("Button1", "TextBox9")).value;
        document.getElementById(e.id.replace("Button1", "TextBox7")).value = document.getElementById(e.id.replace("Button1", "TextBox10")).value;
        document.getElementById(e.id.replace("Button1", "TextBox9")).value = tb6;
        document.getElementById(e.id.replace("Button1", "TextBox10")).value = tb7;
    } else {
        alert('這一筆已經有序號了，表示已結算，無法變更405或409。');
    }
    return false;
}