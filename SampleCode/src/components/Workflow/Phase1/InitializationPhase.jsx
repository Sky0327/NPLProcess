import React from 'react'
import {
  Box,
  Card,
  CardContent,
  Typography,
  TextField,
  Button,
  Grid,
  Divider,
  Alert,
} from '@mui/material'
import { Save, NavigateNext, FolderOpen } from '@mui/icons-material'
import useWorkflowStore from '../../../store/workflowStore'
import { PHASES } from '../../../data/workflowConfig'

const InitializationPhase = () => {
  const {
    projectConfig,
    setProjectConfig,
    setActivePhase,
    addLogEntry,
    calculatePhaseProgress,
  } = useWorkflowStore()

  const phase = PHASES[1]
  const progress = calculatePhaseProgress(1)
  const isComplete = progress === 100

  const handleChange = (field) => (event) => {
    setProjectConfig({ [field]: event.target.value })
  }

  const handleSave = () => {
    addLogEntry({
      type: 'success',
      action: '프로젝트 설정 저장',
      details: `보고서명: ${projectConfig.reportName}`,
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
      action: 'Phase 1 완료',
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
          프로젝트 기본 정보를 설정합니다.
        </Typography>
      </Box>

      <Grid container spacing={3}>
        {/* 프로젝트 정보 */}
        <Grid item xs={12} md={6}>
          <Card>
            <CardContent>
              <Typography variant="subtitle1" sx={{ fontWeight: 600, mb: 2 }}>
                프로젝트 정보
              </Typography>

              <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                <TextField
                  label="보고서명"
                  value={projectConfig.reportName}
                  onChange={handleChange('reportName')}
                  fullWidth
                  size="small"
                  placeholder="예: 2024년 1분기 NPL 평가"
                  required
                />

                <TextField
                  label="프로젝트 ID"
                  value={projectConfig.projectId}
                  onChange={handleChange('projectId')}
                  fullWidth
                  size="small"
                  placeholder="예: NPL-2024-001"
                  required
                />

                <TextField
                  label="입력 폴더 경로"
                  value={projectConfig.inputFolderPath}
                  onChange={handleChange('inputFolderPath')}
                  fullWidth
                  size="small"
                  placeholder="예: C:\NPLData\Input"
                  InputProps={{
                    endAdornment: (
                      <FolderOpen
                        sx={{ color: 'text.secondary', cursor: 'pointer' }}
                      />
                    ),
                  }}
                  required
                />
              </Box>
            </CardContent>
          </Card>
        </Grid>

        {/* API 설정 */}
        <Grid item xs={12} md={6}>
          <Card>
            <CardContent>
              <Typography variant="subtitle1" sx={{ fontWeight: 600, mb: 2 }}>
                API 인증 설정
              </Typography>

              <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                <TextField
                  label="API ID"
                  value={projectConfig.apiId}
                  onChange={handleChange('apiId')}
                  fullWidth
                  size="small"
                  placeholder="API 사용자 ID"
                  required
                />

                <TextField
                  label="API Password"
                  type="password"
                  value={projectConfig.apiPassword}
                  onChange={handleChange('apiPassword')}
                  fullWidth
                  size="small"
                  placeholder="API 비밀번호"
                />

                <Alert severity="info" sx={{ fontSize: '0.8rem' }}>
                  API 인증 정보는 외부 데이터 조회 시 사용됩니다.
                </Alert>
              </Box>
            </CardContent>
          </Card>
        </Grid>

        {/* 상태 및 액션 */}
        <Grid item xs={12}>
          <Card>
            <CardContent>
              <Box
                sx={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                }}
              >
                <Box>
                  <Typography variant="body2" sx={{ color: 'text.secondary' }}>
                    설정 진행률
                  </Typography>
                  <Typography variant="h6" sx={{ fontWeight: 600 }}>
                    {progress}% 완료
                  </Typography>
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
                    다음 단계
                  </Button>
                </Box>
              </Box>
            </CardContent>
          </Card>
        </Grid>
      </Grid>
    </Box>
  )
}

export default InitializationPhase
