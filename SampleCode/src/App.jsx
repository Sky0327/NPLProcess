import React from 'react'
import { Routes, Route, Navigate } from 'react-router-dom'
import DashboardLayout from './components/Layout/DashboardLayout'
import UnifiedWorkflow from './components/Workflow/UnifiedWorkflow'

function App() {
  return (
    <DashboardLayout>
      <Routes>
        <Route path="/" element={<UnifiedWorkflow />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </DashboardLayout>
  )
}

export default App
