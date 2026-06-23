// =====================================
// 1. Global variables
// =====================================

let patients = [];
let currentPatient = 0;
let scoreTimer;
let patientListOpen = false;

// =====================================
// 2. Daten laden
// =====================================

// =====================================
// 2.1 Load patients from Excel
// =====================================

function loadPatients() {

    fetch("patients_veritatem.xlsx")

        .then(response => {

            if (!response.ok) {

                throw new Error(
                    "Failed to load Excel file"
                );

            }

            return response.arrayBuffer();

        })

        .then(data => {

            const workbook =
                XLSX.read(data, {
                    type: "array"
                });

            const sheet =
                workbook.Sheets["Patients"];

            if (!sheet) {

                throw new Error(
                    "Sheet 'Patients' not found"
                );

            }

            patients =
                XLSX.utils.sheet_to_json(sheet);
                buildPatientList();
                buildDatasetExplorer();
                updateDashboard();
                console.log(patients[0]);

            console.log(
                "Patients loaded:",
                patients
            );

            if (patients.length === 0) {

                throw new Error(
                    "No patients found"
                );

            }

        })

        .catch(error => {

            console.error(
                "Loading error:",
                error
            );

        });

}


// =====================================
// 2.2 Load patient by ID
// =====================================

function loadPatientById() {

    let input =
        document
            .getElementById(
                "patientIdInput"
            )
            .value
            .trim();

    let index =
        patients.findIndex(
            p => p.patient_id == input
        );

    if (index !== -1) {

        currentPatient = index;

        updateDashboard();

        showToast(
            "✓ Patient loaded",
            "#44c67d"
        );

    }

    else {

        showToast(
            "⚠ Patient ID not found",
            "#ff6b6b"
        );

    }

}

// =====================================
// 3. Dashboard
// =====================================

// =====================================
// 3.1 Update dashboard
// =====================================

function updateDashboard() {

    if (patients.length === 0) return;

    let p = patients[currentPatient];


    // ---------------------
    // Current patient badge
    // ---------------------

    document.getElementById(
        "currentPatientId"
    ).innerHTML =
        p.patient_id;

    document.getElementById(
        "currentPatientInfo"
    ).innerHTML =
        `Age ${p.age} · ${
            Math.floor(p.gest_week)
        }+${
            Math.round(
                (p.gest_week % 1) * 10
            )
        } weeks`;


    updateChips(p);

    let result =
        calculateProbability(p);

    updateRiskCard(
        result.probability,
        result.risk,
        result.tier,
        result.ciLow,
        result.ciHigh
    );

    updateFactors(p);

    updateActions(result.risk);

    updateAuditLog(p);

}

// =====================================
// 3.2 Patient chips
// =====================================


function updateChips(p) {

    const chipsContainer =
        document.getElementById(
            "chipsContainer"
        );

    chipsContainer.innerHTML = `

    <div class="chip">
        Patient ${p.patient_id}
    </div>

    <div class="chip">
        Age ${p.age}
    </div>

    <div class="chip">
        ${Math.floor(p.gest_week)}+
        ${Math.round(
            (p.gest_week % 1) * 10
        )} weeks
    </div>

    <div class="chip">
        BMI ${p.bmi}
    </div>

    <div class="chip">
        Activity
        ${p.activity_level}
    </div>

    <div class="chip">
        Smoking
        ${p.smoking_status}
    </div>

    <div class="chip">
        Parity
        ${p.parity}
    </div>

    <div class="chip">
        Prior GDM
        ${p.prior_gdm ? "✓" : "✗"}
    </div>

    <div class="chip">
        Family history
        ${p.family_history_dm ? "✓" : "✗"}
    </div>

    <div class="chip">
        Fasting glucose
        ${p.fasting_glucose}
        mmol/L
    </div>

    `;

}

// =====================================
// 3.3 Calculate probability
// =====================================

function calculateProbability(p) {

    let score = 0;

    // Alter
    score += p.age * 0.3;

    // BMI
    score += p.bmi * 0.7;

    // Fasting glucose
    score += p.fasting_glucose * 3;

    // Prior GDM
    if (p.prior_gdm)
        score += 25;

    // Family history
    if (p.family_history_dm)
        score += 10;

    // Low activity level
    if (p.activity_level.toLowerCase() === "low")
        score += 7;

    // Former smoker
    if (p.smoking_status.toLowerCase() === "former")
        score += 4;

    // Multiparity
    if (p.parity >= 2)
        score += 3;


    let probability = Math.round(

        100 /
        (
            1 +
            Math.exp(
                -(score - 50) / 10
            )
        )

    );

    let ciLow =
        Math.max(
            0,
            probability - 8
        );

    let ciHigh =
        Math.min(
            100,
            probability + 8
        );


    let risk;
    let tier;


    if (probability >= 70) {

        risk = "HIGH";

        tier = "Tier 3 / 3";

    }

    else if (probability >= 40) {

        risk = "MODERATE";

        tier = "Tier 2 / 3";

    }

    else {

        risk = "LOW";

        tier = "Tier 1 / 3";

    }


    return {

        probability,

        risk,

        tier,

        ciLow,

        ciHigh

    };

}

// =====================================
// 3.4 Risk Card
// =====================================

function updateRiskCard(
    probability,
    risk,
    tier,
    ciLow,
    ciHigh
) {

    animateScore(probability);

    document.querySelector(
        ".high"
    ).innerHTML = risk;

    document.querySelector(
        ".tier"
    ).innerHTML = tier;

    document.querySelector(
        ".conf"
    ).innerHTML =
        `95% CI · ${ciLow}-${ciHigh}%`;

    let circle =
        document.querySelector(
            ".circle"
        );

    if (risk === "HIGH") {

        circle.style.borderColor =
            "#ff9638";

    }

    else if (
        risk === "MODERATE"
    ) {

        circle.style.borderColor =
            "#ffd34d";

    }

    else {

        circle.style.borderColor =
            "#44c67d";

    }


    // Badge oben rechts

    let badge =
        document.querySelector(
            ".badgeRisk"
        );

    badge.innerHTML =
        risk + " RISK";

    if (risk === "HIGH") {

        badge.style.background =
            "#ff9638";

    }

    else if (
        risk === "MODERATE"
    ) {

        badge.style.background =
            "#ffd34d";

    }

    else {

        badge.style.background =
            "#44c67d";

    }

}

// =====================================
// 3.5 Explainability
// =====================================

function updateFactors(p) {

    const factorsContainer =
        document.getElementById(
            "factorsContainer"
        );

    let factors = [];

    factors.push({
        label: `BMI ${p.bmi}`,
        value: (p.bmi / 35 * 0.15).toFixed(2),
        width: Math.min(100, p.bmi * 2.5)
    });

    factors.push({
        label: "Fasting glucose",
        value: (p.fasting_glucose / 6 * 0.12).toFixed(2),
        width: Math.min(100, p.fasting_glucose * 15)
    });

    if (p.prior_gdm) {

        factors.push({
            label: "Prior GDM",
            value: "0.18",
            width: 100
        });

    }

    if (p.family_history_dm) {

        factors.push({
            label: "Family history",
            value: "0.07",
            width: 50
        });

    }

    if (p.age >= 35) {

        factors.push({
            label: "Maternal age",
            value: "0.06",
            width: 45
        });

    }

    if (p.activity_level.toLowerCase() === "low") {

        factors.push({
            label: "Low activity level",
            value: "0.05",
            width: 40
        });

    }

    if (p.smoking_status.toLowerCase() === "former") {

        factors.push({
            label: "Smoking history",
            value: "0.03",
            width: 25
        });

    }

    if (p.parity >= 2) {

        factors.push({
            label: "Parity",
            value: "0.02",
            width: 20
        });

    }

    factorsContainer.innerHTML = "";

    factors.forEach(f => {

        factorsContainer.innerHTML += `

        <div class="row">

            <div class="label">
                ${f.label}
            </div>

            <div class="track">
                <div class="bar orange"
                     style="width:${f.width}%">
                </div>
            </div>

            <div class="value">
                +${f.value}
            </div>

        </div>

        `;

    });

}

// =====================================
// 3.6 Actions
// =====================================

function updateActions(risk) {

    const actionsContainer =
        document.getElementById(
            "actionsContainer"
        );

    if (risk === "HIGH") {

        actionsContainer.innerHTML = `

        <div class="action primary">

            <h3>
                Offer early OGTT
            </h3>

            <p>
                Now — not at 24–28 weeks
            </p>

        </div>

        <div class="action">

            <h3>
                Lifestyle referral
            </h3>

            <p>
                Dietitian + activity plan
            </p>

        </div>

        <div class="action">

            <h3>
                Closer monitoring
            </h3>

            <p>
                2-weekly weight and glucose
            </p>

        </div>

        `;

    }

    else if (risk === "MODERATE") {

        actionsContainer.innerHTML = `

        <div class="action primary">

            <h3>
                Repeat screening
            </h3>

            <p>
                Additional glucose assessment
            </p>

        </div>

        <div class="action">

            <h3>
                Nutrition counselling
            </h3>

            <p>
                Diet recommendations
            </p>

        </div>

        <div class="action">

            <h3>
                Follow-up in 8 weeks
            </h3>

            <p>
                Standard monitoring interval
            </p>

        </div>

        `;

    }

    else {

        actionsContainer.innerHTML = `

        <div class="action primary">

            <h3>
                Routine monitoring
            </h3>

            <p>
                Continue standard prenatal care
            </p>

        </div>

        `;

    }

}

// =====================================
// 3.7 Audit log
// =====================================

function updateAuditLog(p) {

    const auditContainer =
        document.getElementById(
            "auditContainer"
        );

    auditContainer.innerHTML = `

    <div class="auditItem">

        <div>
            09:14 · Model v2.3.1 ·
            Patient ${p.patient_id}
        </div>

        <div class="badge">
            Logged
        </div>

    </div>


    <div class="auditItem">

        <div>
            09:15 · Human oversight acknowledged
        </div>

        <div class="badge">
            Logged
        </div>

    </div>


    <div class="auditItem">

        <div>
            09:16 · Output stored for monitoring
        </div>

        <div class="badge">
            Logged
        </div>

    </div>

    `;

}

// =====================================
// 3.8 Score animation
// =====================================

function animateScore(target) {

    clearInterval(scoreTimer);

    const scoreElement =
        document.querySelector(
            ".score"
        );

    let current =
        parseInt(
            scoreElement.innerHTML
        ) || 0;

    scoreTimer =
        setInterval(() => {

            if (current < target) {

                current++;

            }

            else if (current > target) {

                current--;

            }

            else {

                clearInterval(
                    scoreTimer
                );

            }

            scoreElement.innerHTML =
                current + "%";

        }, 15);

}

// =====================================
// 4. Patientenverwaltung
// =====================================

// =====================================
// 4.1 Build patient list
// =====================================

function buildPatientList(){

    const patientList =
        document.getElementById(
            "patientList"
        );

    patientList.innerHTML = "";

    patients.forEach((p,index)=>{

        patientList.innerHTML += `

        <div class="patientItem
        ${index === currentPatient ? "active" : ""}"
        onclick="selectPatient(${index})">

            ${p.patient_id}

        </div>

        `;

    });

}

// =====================================
// 4.2 Toggle patient list -! Not in use !-
// =====================================

function togglePatientList(){

    const list =
        document.getElementById(
            "patientList"
        );

    patientListOpen =
        !patientListOpen;

    if(patientListOpen){

        list.style.display = "flex";

    }

    else{

        list.style.display = "none";

    }

}

// =====================================
// 4.3 Select patient
// =====================================

function selectPatient(index){

    currentPatient = index;

    updateDashboard();

    buildPatientList();

}

// =====================================
// 4.4 Build dataset explorer
// =====================================
function buildDatasetExplorer(){

    let html = "";

    patients.forEach((p, index) => {

        let deleteButton;

        if (p.custom) {

            deleteButton = `

                <div class="deletePatient"

                    onclick="deletePatient(${index})">

                    ✕

                </div>

            `;

        }

        else {

            deleteButton = "";

        }

        html += `

        <div class="action">

            <div class="datasetRow">

                <div
                    onclick="selectPatientFromDataset(${index})"
                    style="cursor:pointer; flex:1;">

                    <h3>

                        ${p.patient_id}

                    </h3>

                    <p>

                        Age ${p.age} · BMI ${p.bmi}

                    </p>

                </div>

                ${deleteButton}

            </div>

        </div>

        `;

    });

    document
        .getElementById(
            "datasetExplorer"
        )
        .innerHTML =
        html;

}
// =====================================
// 4.5 Select patient from dataset
// =====================================
function selectPatientFromDataset(index){

    currentPatient = index;

    updateDashboard();

    buildPatientList();

    showDashboard();

    showToast(
        "✓ Patient loaded",
        "#44c67d"
    );

}

// =====================================
// 4.6 Add patient
// =====================================
function addPatient(){

    document
        .getElementById(
            "addPatientPanel"
        )
        .style.display =
        "block";

}

function closeAddPatientPanel(){

    document
        .getElementById(
            "addPatientPanel"
        )
        .style.display =
        "none";

}

// =====================================
// 4.7 Save patient
// =====================================

function savePatient(){

    let patient = {

        patient_id:
            document
                .getElementById(
                    "newPatientId"
                )
                .value,

        age:
            Number(
                document
                    .getElementById(
                        "newAge"
                    )
                    .value
            ),

        bmi:
            Number(
                document
                    .getElementById(
                        "newBMI"
                    )
                    .value
            ),

        fasting_glucose:
            Number(
                document
                    .getElementById(
                        "newGlucose"
                    )
                    .value
            ),

        gest_week:11.2,

        activity_level:"moderate",

        smoking_status:"never",

        parity:0,

        prior_gdm:0,

        family_history_dm:0,

        custom:true

    };

    patients.push(
        patient
    );

    buildPatientList();

    buildDatasetExplorer();

    closeAddPatientPanel();

    showToast(
        "✓ Patient added",
        "#44c67d"
    );

}

// =====================================
// 4.8 Delete patient
// =====================================

function deletePatient(index){

    patients.splice(
        index,
        1
    );

    buildPatientList();

    buildDatasetExplorer();

    showToast(
        "✓ Patient deleted",
        "#ff6b6b"
    );

}

// =====================================
// 5. Navigation
// =====================================

function showDashboard(){

    closeMobileMenu();

    hideAllPages();

    document
        .getElementById(
            "dashboardPage"
        )
        .style.display =
        "block";

    document
        .getElementById(
            "patientBadge"
        )
        .style.display =
        "block";

    document
        .getElementById(
            "titleHeader"
        )
        .style.display =
        "block";

    document
        .getElementById(
            "patientSearch"
        )
        .style.display =
        "flex";
    
    document
        .querySelector(
            "#titleHeader h1"
        )
        .innerText =
        "Veritatem";

    document
        .querySelector(
            "#titleHeader .subtitle"
        )
        .innerText =
        "AI-assisted Clinical Decision Support System";

}


function showPatientsPage(){

    closeMobileMenu();

    hideAllPages();

    document
        .getElementById(
            "patientsPage"
        )
        .style.display =
        "block";

    document
        .getElementById(
            "patientBadge"
        )
        .style.display =
        "block";

    document
        .getElementById(
            "titleHeader"
        )
        .style.display =
        "block";

    document
    .getElementById(
        "patientSearch"
    )
    .style.display =
    "flex";

    document.querySelector("#titleHeader h1").innerText =
    "Patient Management";

    document.querySelector("#titleHeader .subtitle").innerText =
    "Dataset explorer and patient builder";

}

function showProjectPage(){

    closeMobileMenu();

    hideAllPages();

    document
        .getElementById(
            "projectPage"
        )
        .style.display =
        "block";

}


function showAboutPanel(){

    closeMobileMenu();

    document
        .getElementById(
            "aboutPanel"
        )
        .style.display =
        "block";

}

function closeAboutPanel(){

    document
        .getElementById(
            "aboutPanel"
        )
        .style.display =
        "none";

}

function hideAllPages(){

    document
        .getElementById(
            "dashboardPage"
        )
        .style.display =
        "none";

    document
        .getElementById(
            "patientsPage"
        )
        .style.display =
        "none";

    document
        .getElementById(
            "projectPage"
        )
        .style.display =
        "none";

    document
        .getElementById(
            "patientBadge"
        )
        .style.display =
        "none";

    document
        .getElementById(
            "titleHeader"
        )
        .style.display =
        "none";

    document
        .getElementById(
            "patientSearch"
        )
        .style.display =
        "none";

}


// =====================================
// 6. UI-Komponenten
// =====================================

// =====================================
// 6.1 toast meldung
// =====================================

function showToast(message, color) {

    let toast =
        document.getElementById(
            "toast"
        );

    toast.innerHTML =
        message;

    toast.style.background =
        color;

    toast.style.opacity =
        "1";

    setTimeout(() => {

        toast.style.opacity =
            "0";

    }, 2500);

}

// =====================================
// 6.2 AboutPanel
// =====================================


function toggleAboutPanel() {

    const panel =
        document.getElementById(
            "aboutPanel"
        );

    if (panel.style.display === "block") {

        panel.style.display = "none";

    }

    else {

        panel.style.display = "block";

    }

}

// =====================================
// 6.3 ThemeSwich
// =====================================

function toggleTheme(){

    document.body.classList.toggle("dark");

}

// =====================================
// 7. Event Listener
// =====================================

document
    .getElementById(
        "patientIdInput"
    )
    .addEventListener(
        "keypress",
        function(event){

            if(
                event.key === "Enter"
            ){

                loadPatientById();

            }

        }
    );

// =====================================
// 8. Open Documents
// =====================================

function openBusinessPlan(){

    window.open(
        "docs/Veritatem_Business_Plan.pdf",
        "_blank"
    );

}

function openPresentation(){

    window.open(
        "docs/Veritatem_Presentation.pdf",
        "_blank"
    );

}

function openSummary(){

    window.open(
        "docs/Veritatem_Summary.pdf",
        "_blank"
    );

}

function openExcel(){

    window.open(
        "docs/patients_veritatem.pdf",
        "_blank"
    );

}



// =====================================
// 9. Responsive
// =====================================

function toggleMobileMenu(){

    document
        .querySelector(
            ".sidebar"
        )
        .classList
        .toggle(
            "open"
        );

}

function closeMobileMenu(){

    document
        .querySelector(
            ".sidebar"
        )
        .classList
        .remove(
            "open"
        );

}



// =====================================
// 10. Start application
// =====================================

loadPatients();
