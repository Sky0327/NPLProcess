import React from 'react'
import { Routes, Route } from 'react-router-dom'
import Layout from './components/Layout/Layout'
import Home from './pages/Home/Home'
import SmartNPL1 from './pages/SmartNPL1/SmartNPL1'
import SmartNPL2 from './pages/SmartNPL2/SmartNPL2'

function App() {
  return (
    <Layout>
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/smart-npl1" element={<SmartNPL1 />} />
        <Route path="/smart-npl2" element={<SmartNPL2 />} />
      </Routes>
    </Layout>
  )
}

export default App

