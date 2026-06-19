// =====================================
// Global variables
// =====================================

let patients = [];
let currentPatient = 0;
let scoreTimer;


// =====================================
// Load patients from Excel
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

            updateDashboard();

        })

        .catch(error => {

            console.error(
                "Loading error:",
                error
            );

        });

}


// =====================================
// Load patient by ID
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

    }

    else {

        alert(
            "Patient not found"
        );

    }

}


// =====================================
// Build patient list
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
// Toggle patient list
// =====================================

let patientListOpen = false;

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
// Select patient
// =====================================

function selectPatient(index){

    currentPatient = index;

    updateDashboard();

    buildPatientList();

}


// =====================================
// Update dashboard
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
// Patient chips
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
// Calculate probability
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
// Risk Card
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
// Explainability
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
// Actions
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
// Audit log
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
// Score animation
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
// AboutPanel
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
// ThemeSwich
// =====================================

function toggleTheme() {

    let theme =
        document.getElementById(
            "themeStylesheet"
        );

    if (
        theme.getAttribute("href")
        === "style.css"
    ) {

        theme.setAttribute(
            "href",
            "dark-style.css"
        );

    }

    else {

        theme.setAttribute(
            "href",
            "style.css"
        );

    }

}

// =====================================
// Start application
// =====================================

loadPatients();