let a;
let date;
let time;
const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'};
//We added setInterval because we need to update the time every second.
setInterval(() => {
    a = new Date();
    date = a.toLocaleDateString(undefined,options);
    time = a.getHours()+":"+a.getMinutes()+":"+a.getSeconds();
    document.getElementById('time').innerHTML = time+" <br> on "+date;
}, 1000);
//We need to update every second, so interval = 1000 milliseconds

let click = 0;
function clicked(){
    click++
    if(click%2==0){
        document.body.style.backgroundColor = "white";
        document.getElementById('btn').innerHTML = "Dark Mode";
        document.body.style.color = "darkblue";
        tableBody.style.color = "black";
        tableHeading.style.color = "black";
    }
    else{
        document.body.style.backgroundColor = "darkblue";
        document.getElementById('btn').innerHTML = "Light Mode";
        document.body.style.color = "purple";
        tableBody.style.color = "white";
        tableHeading.style.color = "white";
    }
}
