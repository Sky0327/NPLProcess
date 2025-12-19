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
} from '@mui/material'
import {
  Search,
  CheckCircle,
  RadioButtonUnchecked,
} from '@mui/icons-material'
import { apiService } from '../../services/apiService'

const inquiryProcesses = [
  { id: 'registry', name: '등기조회', api: 'getRegistryInfo' },
  { id: 'landPrice', name: '공시지가조회', api: 'getLandPrice' },
  { id: 'auction', name: '법원경매조회', api: 'getCourtAuction' },
  { id: 'infocareStats', name: '인포케어 통계', api: 'getInfocareStats' },
  { id: 'infocareIntegrated', name: '인포케어 통합', api: 'getInfocareIntegrated' },
  { id: 'realEstatePrice', name: '실거래가조회_국토', api: 'getRealEstatePrice' },
  { id: 'valuemapPrice', name: '실거래가조회_밸류맵', api: 'getValuemapPrice' },
]

const DataInquiryStep = ({
  projectData,
  inquiryResults,
  setInquiryResults,
  onNext,
  onBack,
}) => {
  const [loading, setLoading] = useState({})
  const [errors, setErrors] = useState({})

  const handleInquiry = async (process) => {
    setLoading({ ...loading, [process.id]: true })
    setErrors({ ...errors, [process.id]: null })

    try {
      const result = await apiService[process.api](projectData.inputData)
      if (result.success) {
        setInquiryResults({
          ...inquiryResults,
          [process.id]: result.data,
        })
      } else {
        setErrors({ ...errors, [process.id]: '조회 실패' })
      }
    } catch (error) {
      setErrors({ ...errors, [process.id]: error.message })
    } finally {
      setLoading({ ...loading, [process.id]: false })
    }
  }

  const handleAllInquiry = async () => {
    for (const process of inquiryProcesses) {
      await handleInquiry(process)
    }
  }

  const allCompleted = inquiryProcesses.every(
    (process) => inquiryResults[process.id]
  )

  return (
    <Card>
      <CardContent>
        <Typography variant="h5" component="h2" gutterBottom>
          병렬 데이터 조회
        </Typography>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 3 }}>
          다음 조회 프로세스들은 병렬로 실행 가능합니다. 각 프로세스를 개별 실행하거나
          모두 실행할 수 있습니다.
        </Typography>

        <Box sx={{ mb: 3, display: 'flex', gap: 2 }}>
          <Button
            variant="contained"
            onClick={handleAllInquiry}
            disabled={Object.values(loading).some((v) => v)}
          >
            전체 조회 실행
          </Button>
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
          {inquiryProcesses.map((process) => {
            const isCompleted = !!inquiryResults[process.id]
            const isLoading = loading[process.id]
            const hasError = errors[process.id]

            return (
              <Grid item xs={12} sm={6} md={4} key={process.id}>
                <Card
                  variant="outlined"
                  sx={{
                    height: '100%',
                    borderColor: isCompleted ? 'success.main' : 'divider',
                  }}
                >
                  <CardContent>
                    <Box
                      sx={{
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center',
                        mb: 2,
                      }}
                    >
                      <Typography variant="h6">{process.name}</Typography>
                      {isCompleted ? (
                        <CheckCircle color="success" />
                      ) : (
                        <RadioButtonUnchecked color="disabled" />
                      )}
                    </Box>

                    {hasError && (
                      <Alert severity="error" sx={{ mb: 2 }}>
                        {hasError}
                      </Alert>
                    )}

                    {isCompleted && (
                      <Chip
                        label={`${inquiryResults[process.id]?.length || 0}건 조회 완료`}
                        color="success"
                        size="small"
                        sx={{ mb: 2 }}
                      />
                    )}

                    <Button
                      fullWidth
                      variant={isCompleted ? 'outlined' : 'contained'}
                      startIcon={isLoading ? <CircularProgress size={20} /> : <Search />}
                      onClick={() => handleInquiry(process)}
                      disabled={isLoading}
                    >
                      {isCompleted ? '재조회' : '조회 실행'}
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

export default DataInquiryStep

