import React from 'react'
import { Box } from '@mui/material'
import useWorkflowStore from '../../store/workflowStore'
import QuickStats from '../Dashboard/QuickStats'
import InitializationPhase from './Phase1/InitializationPhase'
import ParallelQueriesPhase from './Phase2/ParallelQueriesPhase'
import IntermediatePhase from './Phase3/IntermediatePhase'
import ReportPhase from './Phase4/ReportPhase'
import FinalProcessingPhase from './Phase5/FinalProcessingPhase'

const UnifiedWorkflow = () => {
  const { activePhase } = useWorkflowStore()

  const renderPhase = () => {
    switch (activePhase) {
      case 1:
        return <InitializationPhase />
      case 2:
        return <ParallelQueriesPhase />
      case 3:
        return <IntermediatePhase />
      case 4:
        return <ReportPhase />
      case 5:
        return <FinalProcessingPhase />
      default:
        return <InitializationPhase />
    }
  }

  return (
    <Box>
      <QuickStats />
      {renderPhase()}
    </Box>
  )
}

export default UnifiedWorkflow
