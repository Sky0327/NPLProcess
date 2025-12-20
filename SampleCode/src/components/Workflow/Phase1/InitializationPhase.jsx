import React from 'react'
import {
  Box,
  Card,
  CardContent,
  Typography,
  Button,
  Alert,
  Chip,
} from '@mui/material'
import {
  Save,
  NavigateNext,
  Info,
} from '@mui/icons-material'
import useWorkflowStore from '../../../store/workflowStore'
import { PHASES } from '../../../data/workflowConfig'
import CoverSheetSection from './CoverSheetSection'

const InitializationPhase = () => {
  const {
    projectConfig,
    coverData,
    setActivePhase,
    addLogEntry,
    calculatePhaseProgress,
  } = useWorkflowStore()

  const phase = PHASES[1]
  const progress = calculatePhaseProgress(1)
  const isComplete = progress === 100

  const handleSave = () => {
    addLogEntry({
      type: 'success',
      action: '기본 설정 저장',
      details: `보고서명: ${projectConfig.reportName}, 차주: ${coverData.borrowerName}`,
      phase: 1,
    })
  }

  const handleNext = () => {
    if (!isComplete) {
      addLogEntry({
        type: 'warning',
        action: '설정 미완료',
        details: '모든 필수 항목을 입력해주세요.',
        phase: 1,
      })
      return
    }

    addLogEntry({
      type: 'info',
      action: '기본값 설정 완료',
      details: '데이터 조회 단계로 이동',
      phase: 1,
    })
    setActivePhase(2)
  }

  return (
    <Box>
      {/* Header */}
      <Box sx={{ mb: 3 }}>
        <Typography
          variant="h5"
          sx={{
            fontWeight: 600,
            color: phase.color,
            display: 'flex',
            alignItems: 'center',
            gap: 1,
          }}
        >
          <Box
            sx={{
              width: 8,
              height: 24,
              bgcolor: phase.color,
              borderRadius: 1,
            }}
          />
          {phase.fullName}
        </Typography>
        <Typography variant="body2" sx={{ color: 'text.secondary', mt: 0.5 }}>
          데이터 조회에 필요한 기본 정보를 입력합니다.
        </Typography>
      </Box>

      {/* Info Alert */}
      <Alert severity="info" sx={{ mb: 3 }} icon={<Info />}>
        <Typography variant="body2">
          아래 정보를 입력한 후, <strong>다음 단계</strong>에서 외부 데이터를 조회하면
          채권/담보물 상세 정보가 자동으로 입력됩니다.
        </Typography>
      </Alert>

      {/* Cover Sheet Section */}
      <CoverSheetSection />

      {/* Status and Actions */}
      <Card sx={{ mt: 3 }}>
        <CardContent>
          <Box
            sx={{
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              flexWrap: 'wrap',
              gap: 2,
            }}
          >
            <Box>
              <Typography variant="body2" sx={{ color: 'text.secondary' }}>
                설정 진행률
              </Typography>
              <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
                <Typography variant="h6" sx={{ fontWeight: 600 }}>
                  {progress}% 완료
                </Typography>
                {!isComplete && (
                  <Chip
                    label="필수: 보고서명, 프로젝트ID, 매각기관, 차주명, 소재지"
                    size="small"
                    color="warning"
                    variant="outlined"
                  />
                )}
                {isComplete && (
                  <Chip
                    label="입력 완료"
                    size="small"
                    color="success"
                  />
                )}
              </Box>
            </Box>

            <Box sx={{ display: 'flex', gap: 2 }}>
              <Button
                variant="outlined"
                startIcon={<Save />}
                onClick={handleSave}
              >
                저장
              </Button>
              <Button
                variant="contained"
                endIcon={<NavigateNext />}
                onClick={handleNext}
                disabled={!isComplete}
              >
                데이터 조회로 이동
              </Button>
            </Box>
          </Box>
        </CardContent>
      </Card>
    </Box>
  )
}

export default InitializationPhase
