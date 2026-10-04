import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import OnboardingTour, { STORAGE_KEY } from './OnboardingTour'
import './App.css'

function getWeekDisplay(value) {
  return `${Math.floor(value)}+${Math.round((value % 1) * 10)} weeks`
}

function calculateProbability(patient) {
  let score = 0

  score += patient.age * 0.3
  score += patient.bmi * 0.7
  score += patient.fasting_glucose * 3

  if (patient.prior_gdm) score += 25
  if (patient.family_history_dm) score += 10
  if (patient.activity_level?.toLowerCase() === 'low') score += 7
  if (patient.smoking_status?.toLowerCase() === 'former') score += 4
  if (patient.parity >= 2) score += 3

  const probability = Math.round(100 / (1 + Math.exp(-(score - 50) / 10)))
  const ciLow = Math.max(0, probability - 8)
  const ciHigh = Math.min(100, probability + 8)

  let risk = 'LOW'
  let tier = 'Tier 1 / 3'

  if (probability >= 70) {
    risk = 'HIGH'
    tier = 'Tier 3 / 3'
  } else if (probability >= 40) {
    risk = 'MODERATE'
    tier = 'Tier 2 / 3'
  }

  return { probability, risk, tier, ciLow, ciHigh }
}

function buildFactors(patient) {
  const factors = []

  factors.push({
    label: `BMI ${patient.bmi}`,
    value: (patient.bmi / 35 * 0.15).toFixed(2),
    width: Math.min(100, patient.bmi * 2.5),
  })

  factors.push({
    label: 'Fasting glucose',
    value: (patient.fasting_glucose / 6 * 0.12).toFixed(2),
    width: Math.min(100, patient.fasting_glucose * 15),
  })

  if (patient.prior_gdm) {
    factors.push({ label: 'Prior GDM', value: '0.18', width: 100 })
  }

  if (patient.family_history_dm) {
    factors.push({ label: 'Family history', value: '0.07', width: 50 })
  }

  if (patient.age >= 35) {
    factors.push({ label: 'Maternal age', value: '0.06', width: 45 })
  }

  if (patient.activity_level?.toLowerCase() === 'low') {
    factors.push({ label: 'Low activity level', value: '0.05', width: 40 })
  }

  if (patient.smoking_status?.toLowerCase() === 'former') {
    factors.push({ label: 'Smoking history', value: '0.03', width: 25 })
  }

  if (patient.parity >= 2) {
    factors.push({ label: 'Parity', value: '0.02', width: 20 })
  }

  return factors
}

function buildActions(risk) {
  if (risk === 'HIGH') {
    return [
      { title: 'Offer early OGTT', subtitle: 'Now — not at 24–28 weeks', primary: true },
      { title: 'Lifestyle referral', subtitle: 'Dietitian + activity plan', primary: false },
      { title: 'Closer monitoring', subtitle: '2-weekly weight and glucose', primary: false },
    ]
  }

  if (risk === 'MODERATE') {
    return [
      { title: 'Repeat screening', subtitle: 'Additional glucose assessment', primary: true },
      { title: 'Nutrition counselling', subtitle: 'Diet recommendations', primary: false },
      { title: 'Follow-up in 8 weeks', subtitle: 'Standard monitoring interval', primary: false },
    ]
  }

  return [{ title: 'Routine monitoring', subtitle: 'Continue standard prenatal care', primary: true }]
}

function buildAuditLog(patient) {
  return [
    { text: `09:14 · Model v2.3.1 · Patient ${patient.patient_id}`, status: 'Logged' },
    { text: '09:15 · Human oversight acknowledged', status: 'Logged' },
    { text: '09:16 · Output stored for monitoring', status: 'Logged' },
  ]
}

function App() {
  const [patients, setPatients] = useState([])
  const [currentPatientIndex, setCurrentPatientIndex] = useState(0)
  const [page, setPage] = useState('dashboard')
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const [patientIdInput, setPatientIdInput] = useState('')
  const [searchMessage, setSearchMessage] = useState('')
  const [toast, setToast] = useState({ message: '', color: '#44c67d', visible: false })
  const [darkMode, setDarkMode] = useState(false)
  const [showAddPatientPanel, setShowAddPatientPanel] = useState(false)
  const [tourSeen, setTourSeen] = useState(() => {
    try {
      return localStorage.getItem(STORAGE_KEY) === 'true'
    } catch {
      return false
    }
  })
  const [tourOpen, setTourOpen] = useState(false)
  const [tourStep, setTourStep] = useState(0)
  const [newPatient, setNewPatient] = useState({
    patient_id: '',
    age: '',
    bmi: '',
    fasting_glucose: '',
  })

  useEffect(() => {
    const loadPatients = async () => {
      try {
        const response = await fetch('/patients_veritatem.xlsx')

        if (!response.ok) {
          throw new Error('Failed to load Excel file')
        }

        const data = await response.arrayBuffer()
        const workbook = XLSX.read(data, { type: 'array' })
        const sheet = workbook.Sheets['Patients']

        if (!sheet) {
          throw new Error("Sheet 'Patients' not found")
        }

        const loadedPatients = XLSX.utils.sheet_to_json(sheet)

        if (loadedPatients.length === 0) {
          throw new Error('No patients found')
        }

        setPatients(loadedPatients)
      } catch (error) {
        console.error('Loading error:', error)
      }
    }

    loadPatients()
  }, [])

  useEffect(() => {
    document.body.classList.toggle('dark', darkMode)
  }, [darkMode])

  useEffect(() => {
    if (!tourSeen) {
      setTourOpen(true)
    }
  }, [tourSeen])

  useEffect(() => {
    if (!tourOpen) {
      setSidebarOpen(false)
      return
    }

    setSidebarOpen(tourStep >= 2)
  }, [tourOpen, tourStep])

  useEffect(() => {
    if (!toast.visible) return undefined

    const timer = window.setTimeout(() => {
      setToast((previous) => ({ ...previous, visible: false }))
    }, 2500)

    return () => window.clearTimeout(timer)
  }, [toast.visible])

  const currentPatient = patients[currentPatientIndex] || null
  const dashboardResult = useMemo(
    () => (currentPatient ? calculateProbability(currentPatient) : null),
    [currentPatient],
  )
  const factorRows = useMemo(
    () => (currentPatient ? buildFactors(currentPatient) : []),
    [currentPatient],
  )
  const actionCards = useMemo(
    () => (dashboardResult ? buildActions(dashboardResult.risk) : []),
    [dashboardResult],
  )
  const auditEntries = useMemo(
    () => (currentPatient ? buildAuditLog(currentPatient) : []),
    [currentPatient],
  )

  const openDocument = (path) => {
    window.open(path, '_blank', 'noopener,noreferrer')
  }

  const showToast = (message, color = '#44c67d') => {
    setToast({ message, color, visible: true })
  }

  const handlePatientLoad = (event) => {
    event.preventDefault()

    if (patients.length === 0) return

    const patientId = patientIdInput.trim()
    const index = patients.findIndex((patient) => patient.patient_id == patientId)

    if (index !== -1) {
      setCurrentPatientIndex(index)
      setPage('dashboard')
      setSearchMessage('')
      showToast('✓ Patient loaded', '#44c67d')
      return
    }

    setSearchMessage('Patient ID not found')
    showToast('⚠ Patient ID not found', '#ff6b6b')
  }

  const selectPatient = (index) => {
    setCurrentPatientIndex(index)
    setPage('dashboard')
    setSidebarOpen(false)
  }

  const handleDeletePatient = (index) => {
    const updatedPatients = patients.filter((_, patientIndex) => patientIndex !== index)

    setPatients(updatedPatients)
    setCurrentPatientIndex((previousIndex) => {
      if (previousIndex > index) return previousIndex - 1
      if (previousIndex === index) return Math.max(0, updatedPatients.length - 1)
      return previousIndex
    })

    showToast('✓ Patient deleted', '#ff6b6b')
  }

  const handleSavePatient = () => {
    const patientId = newPatient.patient_id.trim()
    const age = Number(newPatient.age)
    const bmi = Number(newPatient.bmi)
    const fastingGlucose = Number(newPatient.fasting_glucose)

    if (!patientId || Number.isNaN(age) || Number.isNaN(bmi) || Number.isNaN(fastingGlucose)) {
      showToast('⚠ Please complete all fields', '#ff6b6b')
      return
    }

    const nextPatient = {
      patient_id: patientId,
      age,
      bmi,
      fasting_glucose: fastingGlucose,
      gest_week: 11.2,
      activity_level: 'moderate',
      smoking_status: 'never',
      parity: 0,
      prior_gdm: false,
      family_history_dm: false,
      custom: true,
    }

    setPatients((previous) => [...previous, nextPatient])
    setCurrentPatientIndex(patients.length)
    setShowAddPatientPanel(false)
    setNewPatient({ patient_id: '', age: '', bmi: '', fasting_glucose: '' })
    showToast('✓ Patient added', '#44c67d')
  }

  const pageTitle = page === 'dashboard' ? 'Veritatem' : 'Patient Management'
  const pageSubtitle =
    page === 'dashboard'
      ? 'AI-assisted Clinical Decision Support System'
      : 'Dataset explorer and patient builder'

  const riskBadgeStyle =
    dashboardResult?.risk === 'HIGH'
      ? { background: '#ff9638' }
      : dashboardResult?.risk === 'MODERATE'
        ? { background: '#ffd34d' }
        : { background: '#44c67d' }

  const circleBorderColor =
    dashboardResult?.risk === 'HIGH'
      ? '#ff9638'
      : dashboardResult?.risk === 'MODERATE'
        ? '#ffd34d'
        : '#44c67d'

  return (
    <>
      <div className="bg1" />
      <div className="bg2" />
      <div className="bg3" />

      <div className="container">
      <div id="toast" style={{ opacity: toast.visible ? 1 : 0, background: toast.color }}>
        {toast.message}
      </div>

      <aside className={`sidebar ${sidebarOpen ? 'open' : ''}`}>
        <div className="closeSidebar" onClick={() => setSidebarOpen(false)}>
          ✕
        </div>

        <div className="logo" onClick={() => setDarkMode((previous) => !previous)}>
          V
        </div>

        <div className="icon" onClick={() => { setPage('dashboard'); setSidebarOpen(false) }}>
          <img src="/assets/icons/home-icon.svg" className="sidebarIcon" alt="Dashboard" />
        </div>

        <div className="icon" onClick={() => { setPage('patients'); setSidebarOpen(false) }}>
          <img src="/assets/icons/user-icon.svg" className="sidebarIcon" alt="Patients" />
        </div>

        <div id="patientList" style={{ display: 'none' }} />

        <div className="icon" onClick={() => { setPage('project'); setSidebarOpen(false) }}>
          <img src="/assets/icons/project_libary-icon.svg" className="sidebarIcon" alt="Project library" />
        </div>
      </aside>

      <main className="content">
        <header className="header">
          <div className="menuButton" onClick={() => setSidebarOpen((previous) => !previous)}>
            ☰
          </div>

          {page !== 'project' && (
            <div id="titleHeader">
              <h1>{pageTitle}</h1>
              <div className="subtitle">{pageSubtitle}</div>
            </div>
          )}

          {page !== 'project' && (
            <form id="patientSearch" className="patientSearch" onSubmit={handlePatientLoad}>
              <input
                type="text"
                id="patientIdInput"
                placeholder="Enter Patient ID"
                value={patientIdInput}
                onChange={(event) => setPatientIdInput(event.target.value)}
              />

              <button type="submit">Load Patient</button>
            </form>
          )}

          {searchMessage && <div id="searchMessage">{searchMessage}</div>}

          {page !== 'project' && currentPatient && (
            <div className="patientBadge" onClick={() => setPage('patients')}>
              <div className="badgeRisk" style={riskBadgeStyle}>
                {dashboardResult ? `${dashboardResult.risk} RISK` : 'Current patient'}
              </div>

              <div className="badgePatient">{currentPatient.patient_id}</div>

              <div className="badgeInfo">
                Age {currentPatient.age} · {getWeekDisplay(currentPatient.gest_week || 11.2)}
              </div>
            </div>
          )}
        </header>

        <div id="dashboardPage" style={{ display: page === 'dashboard' ? 'block' : 'none' }}>
          <div className="chips" id="chipsContainer">
            {currentPatient && (
              <>
                <div className="chip">Patient {currentPatient.patient_id}</div>
                <div className="chip">Age {currentPatient.age}</div>
                <div className="chip">{getWeekDisplay(currentPatient.gest_week || 11.2)}</div>
                <div className="chip">BMI {currentPatient.bmi}</div>
                <div className="chip">Activity {currentPatient.activity_level}</div>
                <div className="chip">Smoking {currentPatient.smoking_status}</div>
                <div className="chip">Parity {currentPatient.parity}</div>
                <div className="chip">Prior GDM {currentPatient.prior_gdm ? '✓' : '✗'}</div>
                <div className="chip">Family history {currentPatient.family_history_dm ? '✓' : '✗'}</div>
                <div className="chip">
                  Fasting glucose {currentPatient.fasting_glucose} mmol/L
                </div>
              </>
            )}
          </div>

          {currentPatient && dashboardResult && (
            <div className="grid">
              <div className="card">
                <div className="section">Risk Tier</div>

                <div className="circle" style={{ borderColor: circleBorderColor }}>
                  <div className="high">{dashboardResult.risk}</div>
                  <div className="tier">{dashboardResult.tier}</div>
                </div>

                <div className="score">{dashboardResult.probability}%</div>
                <div className="scoreText">Estimated probability</div>
                <div className="conf">95% CI · {dashboardResult.ciLow}-{dashboardResult.ciHigh}%</div>
              </div>

              <div className="card">
                <div className="section">Why this patient</div>
                <div id="factorsContainer">
                  {factorRows.map((factor) => (
                    <div className="row" key={factor.label}>
                      <div className="label">{factor.label}</div>
                      <div className="track">
                        <div className="bar orange" style={{ width: `${factor.width}%` }} />
                      </div>
                      <div className="value">+{factor.value}</div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="card">
                <div className="section">Recommended Actions</div>
                <div id="actionsContainer">
                  {actionCards.map((action) => (
                    <div key={action.title} className={`action ${action.primary ? 'primary' : ''}`}>
                      <h3>{action.title}</h3>
                      <p>{action.subtitle}</p>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          <div className="card auditCard">
            <div className="section">AI Act / MDR Audit Log</div>
            <div id="auditContainer">
              {auditEntries.map((entry) => (
                <div key={entry.text} className="auditItem">
                  <div>{entry.text}</div>
                  <div className="badge">{entry.status}</div>
                </div>
              ))}
            </div>
          </div>
        </div>

        <div id="patientsPage" style={{ display: page === 'patients' ? 'block' : 'none' }}>
          <div className="section">Dataset</div>

          <div id="datasetExplorer">
            {patients.map((patient, index) => (
              <div key={`${patient.patient_id}-${index}`} className="action">
                <div className="datasetRow">
                  <div
                    onClick={() => {
                      setCurrentPatientIndex(index)
                      setPage('dashboard')
                      showToast('✓ Patient loaded', '#44c67d')
                    }}
                    style={{ cursor: 'pointer', flex: 1 }}
                  >
                    <h3>{patient.patient_id}</h3>
                    <p>
                      Age {patient.age} · BMI {patient.bmi}
                    </p>
                  </div>

                  {patient.custom && (
                    <div className="deletePatient" onClick={() => handleDeletePatient(index)}>
                      ✕
                    </div>
                  )}
                </div>
              </div>
            ))}
          </div>

          <div className="action primary" onClick={() => setShowAddPatientPanel(true)}>
            <h3>+ Add Patient</h3>
          </div>
        </div>

        <div id="projectPage" style={{ display: page === 'project' ? 'block' : 'none' }}>
          <h1>Project Library</h1>
          <div className="subtitle">Documents, datasets and mockups</div>

          <br />
          <br />

          <div className="projectInfo">
            <div className="section">About the Project</div>
            <p>
              This dashboard illustrates how the Veritatem project and its AI-assisted Software-as-a-Service (SaaS) clinical decision support system (CDSS) could be visualized. The prototype is based on an external Excel dataset containing three fictional patient records representing low-, moderate-, and high-risk scenarios.
              <br />
              The Patient Manager allows users to switch between patient profiles and create additional demonstration profiles to showcase the functionality of the system. These entries are not stored in the underlying dataset. All generated values, probabilities, and visualizations are illustrative and are not intended for clinical use.
              <br />
              Veritatem explores how AI-assisted decision support could help clinicians identify and manage gestational diabetes at an earlier stage. The system provides recommendations only, while all diagnostic and treatment decisions remain the responsibility of the clinician.
              <br />
              In addition to the technical concept, the project examines regulatory, economic, and ethical aspects of AI in healthcare, including the MDR, AI Act, GDPR, and data acquisition. Further information can be found in the Documents section, which provides access to the business plan, presentation, and executive summary.
            </p>
          </div>

          <div className="card nohoverCard">
            <div className="section">Dataset Structure</div>

            <div className="datasetLayout">
              <div className="datasetBox">
                <table className="datasetTable">
                  <thead>
                    <tr>
                      <th>patient_id</th>
                      <th>age</th>
                      <th>gest_week</th>
                      <th>bmi</th>
                      <th>activity_level</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td>PX-00417</td>
                      <td>22</td>
                      <td>11.2</td>
                      <td>20.8</td>
                      <td>high</td>
                    </tr>
                    <tr>
                      <td>PX-00284</td>
                      <td>29</td>
                      <td>12.1</td>
                      <td>26.1</td>
                      <td>moderate</td>
                    </tr>
                    <tr>
                      <td>PX-00732</td>
                      <td>35</td>
                      <td>11.4</td>
                      <td>33.4</td>
                      <td>low</td>
                    </tr>
                  </tbody>
                </table>

                <div className="datasetSource">
                  Source:
                  <span onClick={() => openDocument('/docs/patients_veritatem.pdf')} className="sourceLink">
                    {' '}
                    patients_veritatem.xlsx
                  </span>
                  <br />
                  Example subset of the dataset. Additional variables are available in the source file.
                </div>
              </div>

              <div className="card nohoverCard documentsBox">
                <div className="section">Documents</div>

                <div className="action documentCard" onClick={() => openDocument('/docs/Veritatem_Business_Plan.pdf')}>
                  <div className="openBadge">Open ↗</div>
                  <h3>Business Plan</h3>
                  <p>Financial model and market analysis</p>
                </div>

                <div className="action documentCard" onClick={() => openDocument('/docs/Veritatem_Presentation.pdf')}>
                  <div className="openBadge">Open ↗</div>
                  <h3>Presentation</h3>
                  <p>Pitch deck slides</p>
                </div>

                <div className="action documentCard" onClick={() => openDocument('/docs/Veritatem_Summary.pdf')}>
                  <div className="openBadge">Open ↗</div>
                  <h3>Summary</h3>
                  <p>Get a broad overview of our project</p>
                </div>
              </div>
            </div>
          </div>

          <div className="card nohoverCard">
            <div className="section">Project Information</div>
            <div className="disclaimerGrid">
              <div>
                <h4>Course</h4>
                <p>26S MDK.I04 Economic Aspects</p>
              </div>
              <div>
                <h4>Lecturer</h4>
                <p>Dipl.-Ing. (FH) MA Michael Freidl</p>
              </div>
              <div>
                <h4>Prepared by</h4>
                <p>
                  Julian Baron,
                  <br />
                  Zan Porenta,
                  <br />
                  Satera Sleinyte
                </p>
              </div>
              <div>
                <h4>Institution</h4>
                <p>Karl-Franzens-Universität Graz</p>
              </div>
              <div>
                <h4>Dashboard Visualization</h4>
                <p>© 2026 Julian Baron</p>
              </div>
            </div>
          </div>
        </div>
      </main>

      <OnboardingTour
        enabled={tourOpen}
        onStepChange={setTourStep}
        onComplete={() => {
          setTourSeen(true)
          setTourOpen(false)
          setSidebarOpen(false)
        }}
        onSkip={() => {
          setTourSeen(true)
          setTourOpen(false)
          setSidebarOpen(false)
        }}
      />

      <div id="aboutPanel" style={{ display: 'none' }}>
        <div className="closeButton" onClick={() => {}}>
          ✕
        </div>

        <h2>Veritatem</h2>
        <p>Clinical Decision Support System</p>
        <p>Prototype version 0.1</p>
        <hr />
        <p>
          Developed by
          <br />
          Julian Baron
        </p>
        <p>University of Graz</p>
        <p>
          Bachelor Programme:
          <br />
          Law, Economics and Society in the Digital Age
        </p>
        <hr />
        <p>For academic and demonstration purposes only.</p>
      </div>

      {showAddPatientPanel && (
        <div id="addPatientPanel" style={{ display: 'block' }}>
          <div className="closeButton" onClick={() => setShowAddPatientPanel(false)}>
            ✕
          </div>

          <h2>Add Patient</h2>
          <br />

          <input
            id="newPatientId"
            placeholder="Patient ID"
            value={newPatient.patient_id}
            onChange={(event) => setNewPatient((previous) => ({ ...previous, patient_id: event.target.value }))}
          />

          <input
            id="newAge"
            placeholder="Age"
            value={newPatient.age}
            onChange={(event) => setNewPatient((previous) => ({ ...previous, age: event.target.value }))}
          />

          <input
            id="newBMI"
            placeholder="BMI"
            value={newPatient.bmi}
            onChange={(event) => setNewPatient((previous) => ({ ...previous, bmi: event.target.value }))}
          />

          <input
            id="newGlucose"
            placeholder="Fasting glucose"
            value={newPatient.fasting_glucose}
            onChange={(event) =>
              setNewPatient((previous) => ({ ...previous, fasting_glucose: event.target.value }))
            }
          />

          <button type="button" onClick={handleSavePatient}>
            Save Patient
          </button>
        </div>
      )}
      </div>
    </>
  )
}

export default App
