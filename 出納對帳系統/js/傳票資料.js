/*function check(e) {
    var i = 0;
    while (document.getElementById(e.id.replace("Button4", "GridView1_CheckBox1_" + i.toString())) !== null) {
        var checkbox = document.getElementById(e.id.replace("Button4", "GridView1_CheckBox1_" + i.toString()))
        if (!checkbox.disabled) {
            checkbox.checked = !checkbox.checked;
        }
        i++;
    }
    return false;
}*/

function totaiwancalendar1() {
    var Calendar1 = document.getElementById("ContentPlaceHolder1_Calendar1");
    var before = parseInt(Calendar1.value.substr(0, 4));
    var after = before - 1911;
    Calendar1.Value = Calendar1.Value.Replace(before, after);
    return false;
}

function totaiwancalendar2() {
    var Calendar2 = document.getElementById("ContentPlaceHolder1_Calendar2");
    var before = parseInt(Calendar2.value.substr(0, 4));
    var after = before - 1911;
    Calendar2.Value = Calendar2.Value.Replace(before, after);
    return false;
}