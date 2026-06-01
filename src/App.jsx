import ExportButton from './ExportButton'
import { aprilData } from './aprilData'

function App() {
  return (
    <div className="min-h-screen flex items-center justify-center ">
      <div className="text-center space-y-4">
        <h1 className="text-xl font-semibold text-gray-800">
          Anjuman Wazifa-e-Sadat Wa Momineen
        </h1>
        <p className="text-gray-500">April 2026 Report</p>
        {/* Static data use */}
        <ExportButton data={aprilData} fileName="April-2026.xlsx" />

        {/* API se data aane par bas data prop replace karo:
            <ExportButton data={apiData} fileName="April-2026.xlsx" />
        */}
      </div>
    </div>
  )
}

export default App
