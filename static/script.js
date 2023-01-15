function goTo(){
    value = document.getElementById('select').value
    if(value == "one"){
        location.href="/oneDay/";
    }
    else if(value == "seven"){
        location.href="/sevenDays/";
    }
    else if(value == "thirty"){
        location.href="/thirtyDays/";
    }
}

function backToMainPage(){
    location.href ="/";
}
