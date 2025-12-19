import React, { useState } from 'react'
import {
  Box,
  Typography,
  Stepper,
  Step,
  StepLabel,
  Paper,
  Grid,
} from '@mui/material'
import InitializationStep from '../../components/SmartNPL1/InitializationStep'
import DataInquiryStep from '../../components/SmartNPL1/DataInquiryStep'
import IntermediateProcessingStep from '../../components/SmartNPL1/IntermediateProcessingStep'
import ReportGenerationStep from '../../components/SmartNPL1/ReportGenerationStep'

const steps = [
  '초기화 및 설정',
  '병렬 데이터 조회',
  '중간 처리',
  '리포트 생성',
]

const SmartNPL1 = () => {
  const [activeStep, setActiveStep] = useState(0)
  const [projectData, setProjectData] = useState({
    reportName: '',
    projectId: '',
    apiSettings: {},
    inputData: [],
  })
  const [inquiryResults, setInquiryResults] = useState({})
  const [processingResults, setProcessingResults] = useState({})

  const handleNext = () => {
    setActiveStep((prevActiveStep) => prevActiveStep + 1)
  }

  const handleBack = () => {
    setActiveStep((prevActiveStep) => prevActiveStep - 1)
  }

  const handleStepChange = (step) => {
    setActiveStep(step)
  }

  const renderStepContent = (step) => {
    switch (step) {
      case 0:
        return (
          <InitializationStep
            projectData={projectData}
            setProjectData={setProjectData}
            onNext={handleNext}
          />
        )
      case 1:
        return (
          <DataInquiryStep
            projectData={projectData}
            inquiryResults={inquiryResults}
            setInquiryResults={setInquiryResults}
            onNext={handleNext}
            onBack={handleBack}
          />
        )
      case 2:
        return (
          <IntermediateProcessingStep
            inquiryResults={inquiryResults}
            processingResults={processingResults}
            setProcessingResults={setProcessingResults}
            onNext={handleNext}
            onBack={handleBack}
          />
        )
      case 3:
        return (
          <ReportGenerationStep
            projectData={projectData}
            inquiryResults={inquiryResults}
            processingResults={processingResults}
            onBack={handleBack}
          />
        )
      default:
        return null
    }
  }

  return (
    <Box>
      <Typography variant="h4" component="h1" gutterBottom sx={{ mb: 4, fontWeight: 700 }}>
        Smart_NPL1 - 데이터 수집 및 처리
      </Typography>

      <Paper sx={{ p: 3, mb: 4 }}>
        <Stepper activeStep={activeStep} alternativeLabel>
          {steps.map((label, index) => (
            <Step key={label}>
              <StepLabel
                onClick={() => handleStepChange(index)}
                sx={{ cursor: 'pointer' }}
              >
                {label}
              </StepLabel>
            </Step>
          ))}
        </Stepper>
      </Paper>

      <Grid container spacing={3}>
        <Grid item xs={12}>
          {renderStepContent(activeStep)}
        </Grid>
      </Grid>
    </Box>
  )
}

export default SmartNPL1

