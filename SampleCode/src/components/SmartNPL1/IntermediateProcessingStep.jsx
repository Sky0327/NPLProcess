import React, { useState } from 'react'
import {
  Box,
  Button,
  Card,
  CardContent,
  Typography,
  Grid,
  CircularProgress,
  Chip,
  Alert,
  Divider,
} from '@mui/material'
import {
  PlayArrow,
  CheckCircle,
  RadioButtonUnchecked,
} from '@mui/icons-material'
import { apiService } from '../../services/apiService'

const processingSteps = [
  {
    id: 'kbPrice',
    name: 'KB시세조회',
    description: '등기조회 + 공시지가조회 결과 기반',
    dependencies: ['registry', 'landPrice'],
    api: 'getKBPrice',
  },
  {
    id: 'infocareCases',
    name: '인포케어 사례',
    description: '인포케어 통합 결과 기반',
    dependencies: ['infocareIntegrated'],
    api: 'getInfocareCases',
  },
  {
    id: 'distanceGookto',
    name: '거리계산_국토',
    description: '실거래가조회_국토 결과 기반',
    dependencies: ['realEstatePrice'],
    api: 'calculateDistance',
  },
  {
    id: 'distanceValuemap',
    name: '거리계산_밸류맵',
    description: '실거래가조회_밸류맵 결과 기반',
    dependencies: ['valuemapPrice'],
    api: 'calculateDistanceValuemap',
  },
]

const IntermediateProcessingStep = ({
  inquiryResults,
  processingResults,
  setProcessingResults,
  onNext,
  onBack,
}) => {
  const [loading, setLoading] = useState({})
  const [errors, setErrors] = useState({})

  const checkDependencies = (step) => {
    return step.dependencies.every((dep) => inquiryResults[dep])
  }

  const handleProcessing = async (step) => {
    if (!checkDependencies(step)) {
      alert(`필수 데이터가 준비되지 않았습니다: ${step.dependencies.join(', ')}`)
      return
    }

    setLoading({ ...loading, [step.id]: true })
    setErrors({ ...errors, [step.id]: null })

    try {
      let result
      if (step.id === 'kbPrice') {
        result = await apiService[step.api](
          inquiryResults.registry,
          inquiryResults.landPrice
        )
      } else if (step.id === 'infocareCases') {
        result = await apiService[step.api](inquiryResults.infocareIntegrated)
      } else if (step.id === 'distanceGookto') {
        result = await apiService[step.api](inquiryResults.realEstatePrice)
      } else if (step.id === 'distanceValuemap') {
        result = await apiService[step.api](inquiryResults.valuemapPrice)
      }

      if (result.success) {
        setProcessingResults({
          ...processingResults,
          [step.id]: result.data,
        })
      } else {
        setErrors({ ...errors, [step.id]: '처리 실패' })
      }
    } catch (error) {
      setErrors({ ...errors, [step.id]: error.message })
    } finally {
      setLoading({ ...loading, [step.id]: false })
    }
  }

  const allCompleted = processingSteps.every(
    (step) => processingResults[step.id] || !checkDependencies(step)
  )

  return (
    <Card>
      <CardContent>
        <Typography variant="h5" component="h2" gutterBottom>
          중간 처리
        </Typography>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 3 }}>
          조회 결과를 기반으로 추가 처리를 수행합니다. 각 단계는 이전 단계의 결과에
          의존하므로 순차적으로 실행됩니다.
        </Typography>

        <Box sx={{ mb: 3, display: 'flex', gap: 2 }}>
          {allCompleted && (
            <Button variant="contained" color="success" onClick={onNext}>
              다음 단계로
            </Button>
          )}
          <Button variant="outlined" onClick={onBack}>
            이전
          </Button>
        </Box>

        <Grid container spacing={2}>
          {processingSteps.map((step) => {
            const isCompleted = !!processingResults[step.id]
            const isLoading = loading[step.id]
            const hasError = errors[step.id]
            const canExecute = checkDependencies(step)

            return (
              <Grid item xs={12} sm={6} key={step.id}>
                <Card
                  variant="outlined"
                  sx={{
                    height: '100%',
                    borderColor: isCompleted
                      ? 'success.main'
                      : canExecute
                      ? 'primary.main'
                      : 'divider',
                    opacity: canExecute ? 1 : 0.6,
                  }}
                >
                  <CardContent>
                    <Box
                      sx={{
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center',
                        mb: 1,
                      }}
                    >
                      <Typography variant="h6">{step.name}</Typography>
                      {isCompleted ? (
                        <CheckCircle color="success" />
                      ) : (
                        <RadioButtonUnchecked color="disabled" />
                      )}
                    </Box>

                    <Typography
                      variant="body2"
                      color="text.secondary"
                      sx={{ mb: 2 }}
                    >
                      {step.description}
                    </Typography>

                    {!canExecute && (
                      <Alert severity="warning" sx={{ mb: 2 }}>
                        필수 데이터: {step.dependencies.join(', ')}
                      </Alert>
                    )}

                    {hasError && (
                      <Alert severity="error" sx={{ mb: 2 }}>
                        {hasError}
                      </Alert>
                    )}

                    {isCompleted && (
                      <Chip
                        label={`${processingResults[step.id]?.length || 0}건 처리 완료`}
                        color="success"
                        size="small"
                        sx={{ mb: 2 }}
                      />
                    )}

                    <Button
                      fullWidth
                      variant={isCompleted ? 'outlined' : 'contained'}
                      startIcon={
                        isLoading ? <CircularProgress size={20} /> : <PlayArrow />
                      }
                      onClick={() => handleProcessing(step)}
                      disabled={isLoading || !canExecute}
                    >
                      {isCompleted ? '재처리' : '처리 실행'}
                    </Button>
                  </CardContent>
                </Card>
              </Grid>
            )
          })}
        </Grid>
      </CardContent>
    </Card>
  )
}

export default IntermediateProcessingStep

