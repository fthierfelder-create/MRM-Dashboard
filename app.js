let mrmData = [];


// Dateiimport starten

document
.getElementById("fileInput")
.addEventListener("change", function(event) {


    const file = event.target.files[0];


    if (!file) {
        return;
    }


    document.getElementById("importStatus").innerHTML =
        "Datei wird geladen...";


    const reader = new FileReader();



    reader.onload = function(e) {


        const data = new Uint8Array(e.target.result);


        const workbook = XLSX.read(data, {
            type: "array"
        });


        analyseMRM(workbook);


    };


    reader.readAsArrayBuffer(file);


});




function analyseMRM(workbook) {

    mrmData = [];

    let tomCounter = 0;
    let measureCounter = 0;
    let evidenceCounter = 0;
    let isoCounter = 0;
	let b3sCounter = 0;
	let nis2Counter = 0;
	let dsgvoCounter = 0;

    let analysisHTML = "";
let compliance = {

    iso: 0,
    b3s: 0,
    nis2: 0,
    dsgvo: 0

};


let tomOverview = {};

    workbook.SheetNames.forEach(sheetName => {


        if (sheetName.endsWith("_REF")) {


            tomCounter++;


            const sheet = workbook.Sheets[sheetName];


            const rows = XLSX.utils.sheet_to_json(
                sheet,
                {
                    defval: ""
                }
            );
tomOverview[sheetName] = rows.length;

            let tomMeasures = 0;


            rows.forEach(row => {


                tomMeasures++;
                measureCounter++;
				if (
    row["ISO 27001"] &&
    row["ISO 27001"].toString().trim() !== ""
) {
    compliance.iso++;
}


if (
    row["B3S-ID"] &&
    row["B3S-ID"].toString().trim() !== ""
) {
    compliance.b3s++;
}


if (
    row["NIS2"] &&
    row["NIS2"].toString().trim() !== ""
) {
    compliance.nis2++;
}


if (
    row["DSGVO"] &&
    row["DSGVO"].toString().trim() !== ""
) {
    compliance.dsgvo++;
}


               if (
    row["Auditnachweis (EV-ID)"] &&
    row["Auditnachweis (EV-ID)"].toString().trim() !== ""
)
{
    evidenceCounter++;
}


                if (
                    row["ISO 27001"] &&
                    row["ISO 27001"].toString().trim() !== ""
                ) {

                    isoCounter++;

                }
if (
 row["B3S-ID"] &&
 row["B3S-ID"].toString().trim() !== ""
)
{
 b3sCounter++;
}


if (
 row["NIS2"] &&
 row["NIS2"].toString().trim() !== ""
)
{
 nis2Counter++;
}


if (
 row["DSGVO"] &&
 row["DSGVO"].toString().trim() !== ""
)
{
 dsgvoCounter++;
}

                mrmData.push({

                    tom: sheetName,
                    data: row

                });


            });



            analysisHTML += `

            <div class="analysis-item">

                <span class="analysis-title">
                    ${sheetName}
                </span>

                :
                ${tomMeasures}
                Maßnahmen

            </div>

            `;


        }


    });



    document.getElementById("sheetAnalysis").innerHTML =
        analysisHTML;



    updateDashboard(
         tomCounter,
    measureCounter,
    evidenceCounter,
    isoCounter,
    b3sCounter,
    nis2Counter,
    dsgvoCounter
    );
showTomOverview(tomOverview);
showMissingEvidence();


    document.getElementById("importStatus").innerHTML =
        "🟢 MRM erfolgreich analysiert: "
        +
        tomCounter
        +
        " TOM Register erkannt";


}



function updateDashboard(
    tom,
    measures,
    evidence,
    iso,
    b3s,
    nis2,
    dsgvo
)
{
	let rate = 0;


if (measures > 0) {

    rate = Math.round(
        (evidence / measures) * 100
    );

}


document.getElementById("tomCount")
.innerHTML = tom;


document.getElementById("measureCount")
.innerHTML = measures;


document.getElementById("evidenceCount")
.innerHTML = evidence;


document.getElementById("isoCount")
.innerHTML = iso;


document.getElementById("b3sCount")
.innerHTML = b3s;


document.getElementById("nis2Count")
.innerHTML = nis2;


document.getElementById("dsgvoCount")
.innerHTML = dsgvo;

document.getElementById("evidenceRate")
.innerHTML =
rate + " %";
}





function buildTomTable(){


const table =
document.getElementById("tomTable");


table.innerHTML="";



let grouped = {};



mrmData.forEach(item=>{


    if(!grouped[item.tom]){


        grouped[item.tom]={

            count:0,
            document:"",
            responsible:"",
            status:""

        };

    }



    grouped[item.tom].count++;


    grouped[item.tom].document =
    item.document;


    grouped[item.tom].responsible =
    item.responsible;


    grouped[item.tom].status =
    item.status;



});





Object.keys(grouped)
.forEach(tom=>{


const row =
document.createElement("tr");



row.innerHTML = `

<td>${tom}</td>

<td>${grouped[tom].document}</td>

<td>${grouped[tom].count}</td>

<td>${grouped[tom].responsible}</td>

<td>${formatStatus(grouped[tom].status)}</td>

`;



table.appendChild(row);



});



}





function formatStatus(status){


if(!status || status===""){

return "⚪ offen";

}



status=status.toLowerCase();



if(status.includes("umgesetzt")){

return "🟢 umgesetzt";

}



if(status.includes("geplant")){

return "🟡 geplant";

}



return status;



}
function showTomOverview(tomOverview) {


    let html = "";


    Object.keys(tomOverview).forEach(tom => {


        html += `

        <div class="tom-item">

            <button onclick="showTomDetails('${tom}')">

                ${tom}

            </button>

            :

            ${tomOverview[tom]}

            Maßnahmen

        </div>

        `;


    });



    document.getElementById("tomOverviewBox").innerHTML = html;


}
function showTomDetails(selectedTom) {


    let html = "";


    let measures = mrmData.filter(item => 
        item.tom === selectedTom
    );
console.log(
    "Gefundene Felder:",
    Object.keys(measures[0].data).join("\n")
);


    html += `

    <h3>${selectedTom}</h3>

    <p>
    Anzahl Maßnahmen:
    ${measures.length}
    </p>

    `;



    measures.forEach(item => {


        html += `

        <div class="detail-item">

            <div class="detail-title">

                Maßnahme

            </div>


<div class="detail-table">

<table>

<tr>
<td>Dokument</td>
<td>${item.data["Dokument"] || ""}</td>
</tr>


<tr>
<td>Maßnahme</td>
<td>${item.data["Sicherheitsmaßnahme"] || ""}</td>
</tr>


<tr>
<td>Verantwortlich</td>
<td>${item.data["Verantwortlich"] || ""}</td>
</tr>


<tr>
<td>ISO 27001</td>
<td>${item.data["ISO 27001"] || ""}</td>
</tr>


<tr>
<td>B3S</td>
<td>${item.data["B3S-ID"] || ""}</td>
</tr>


<tr>
<td>NIS2</td>
<td>${item.data["NIS2"] || ""}</td>
</tr>


<tr>
<td>DSGVO</td>
<td>${item.data["DSGVO"] || ""}</td>
</tr>


<tr>
<td>Auditnachweis</td>

<td>

${
item.data["Auditnachweis (EV-ID)"]
?
"🟢 " + item.data["Auditnachweis (EV-ID)"]
:
"🔴 kein Nachweis"
}

</td>

</tr>

</table>

</div>


        </div>


        `;


    });



    document.getElementById(
        "tomDetailsBox"
    ).innerHTML = html;


}

function showMissingEvidence() {


    let html = "";


    let missing = mrmData.filter(item =>

        !item.data["Auditnachweis (EV-ID)"]
        ||
        item.data["Auditnachweis (EV-ID)"]
        .toString()
        .trim() === ""

    );



    if (missing.length === 0) {


        html = "🟢 Alle Maßnahmen besitzen einen Nachweis";


    }
    else {


        missing.forEach(item => {


            html += `

            <div class="detail-item">

                <b>${item.tom}</b>
                <br>

                Maßnahme-ID:
                ${item.data["Maßnahme-ID"] || ""}
                <br>

                Dokument:
                ${item.data["Dokument"] || ""}
                <br>

                Verantwortlich:
                ${item.data["Verantwortlich"] || ""}

            </div>

            `;


        });


    }



    document.getElementById(
        "missingEvidenceBox"
    ).innerHTML = html;


}

function openHelp(title, content) {

    document.getElementById("helpTitle").innerHTML = title;

    document.getElementById("helpContent").innerHTML = content;

    document.getElementById("helpModal").style.display = "block";

}


function closeHelp() {

    document.getElementById("helpModal").style.display = "none";

}