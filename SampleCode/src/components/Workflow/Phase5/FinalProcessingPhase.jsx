import React from 'react'
import {
  Box,
  Typography,
  Button,
  Grid,
  Card,
  CardContent,
  LinearProgress,
  Divider,
  Alert,
} from '@mui/material'
import {
  NavigateBefore,
  Description,
  PictureAsPdf,
  FileDownload,
  RestartAlt,
  CheckCircle,
} from '@mui/icons-material'
import useWorkflowStore from '../../../store/workflowStore'
import { PHASES } from '../../../data/workflowConfig'
import TaskCard from '../../Common/TaskCard'

const FinalProcessingPhase = () => {
  const {
    taskResults,
    taskStatus,
    setTaskResult,
    setTaskStatus,
    setActivePhase,
    addLogEntry,
    calculatePhaseProgress,
    calculateOverallProgress,
    resetWorkflow,
  } = useWorkflowStore()

  const phase = PHASES[5]
  const progress = calculatePhaseProgress(5)
  const overallProgress = calculateOverallProgress()
  const isComplete = progress === 100

  const handleTask = async (taskId, taskName, action) => {
    setTaskStatus(taskId, { loading: true, error: null })

    addLogEntry({
      type: 'info',
      action: `${taskName} 시작`,
      phase: 5,
    })

    try {
      await new Promise((resolve) => setTimeout(resolve, 1500))

      setTaskResult(taskId, {
        completedAt: new Date().toISOString(),
        status: 'completed',
      })
      setTaskStatus(taskId, { loading: false, error: null })

      addLogEntry({
        type: 'success',
        action: `${taskName} 완료`,
        phase: 5,
      })
    } catch (error) {
      setTaskStatus(taskId, { loading: false, error: error.message })

      addLogEntry({
        type: 'error',
        action: `${taskName} 오류`,
        details: error.message,
        phase: 5,
      })
    }
  }

  const handleReset = () => {
    addLogEntry({
      type: 'info',
      action: '워크플로우 초기화',
      details: '모든 데이터가 초기화됩니다.',
      phase: 5,
    })
    resetWorkflow()
    setActivePhase(1)
  }

  const finalTasks = [
    {
      id: 'report-물건지',
      name: '[0] 물건지',
      icon: Description,
    },
    {
      id: 'report-채권현황',
      name: '[1] 채권현황',
      icon: Description,
    },
  ]

  const processingTasks = [
    {
      id: 'pdfConversion',
      name: '등본PDF 파일명 변환',
      icon: PictureAsPdf,
    },
    {
      id: 'xlsxExport',
      name: 'XLSX 내보내기',
      icon: FileDownload,
    },
  ]

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
          마지막 리포트 생성 및 파일 내보내기를 수행합니다.
        </Typography>
      </Box>

      {/* Overall Progress */}
      {isComplete && (
        <Alert
          icon={<CheckCircle />}
          severity="success"
          sx={{ mb: 3 }}
        >
          모든 작업이 완료되었습니다! 전체 진행률: {overallProgress}%
        </Alert>
      )}

      {/* Progress */}
      <Card sx={{ mb: 3 }}>
        <CardContent sx={{ py: 2 }}>
          <Box
            sx={{
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              mb: 1,
            }}
          >
            <Typography variant="body2" sx={{ fontWeight: 500 }}>
              Phase 5 진행률
            </Typography>
            <Typography variant="body2" sx={{ color: 'text.secondary' }}>
              {progress}%
            </Typography>
          </Box>
          <LinearProgress
            variant="determinate"
            value={progress}
            sx={{
              height: 8,
              '& .MuiLinearProgress-bar': {
                bgcolor: phase.color,
              },
            }}
          />
        </CardContent>
      </Card>

      {/* Final Reports */}
      <Typography variant="subtitle1" sx={{ fontWeight: 600, mb: 2 }}>
        최종 리포트 생성
      </Typography>
      <Grid container spacing={2} sx={{ mb: 4 }}>
        {finalTasks.map((task) => (
          <Grid item xs={12} sm={6} key={task.id}>
            <TaskCard
              id={task.id}
              name={task.name}
              result={taskResults[task.id]}
              loading={taskStatus[task.id]?.loading}
              error={taskStatus[task.id]?.error}
              onExecute={() => handleTask(task.id, task.name)}
              compact
            />
          </Grid>
        ))}
      </Grid>

      {/* Processing Tasks */}
      <Typography variant="subtitle1" sx={{ fontWeight: 600, mb: 2 }}>
        파일 처리 및 내보내기
      </Typography>
      <Grid container spacing={2} sx={{ mb: 4 }}>
        {processingTasks.map((task) => (
          <Grid item xs={12} sm={6} key={task.id}>
            <TaskCard
              id={task.id}
              name={task.name}
              result={taskResults[task.id]}
              loading={taskStatus[task.id]?.loading}
              error={taskStatus[task.id]?.error}
              onExecute={() => handleTask(task.id, task.name)}
              compact
            />
          </Grid>
        ))}
      </Grid>

      <Divider sx={{ my: 3 }} />

      {/* Actions */}
      <Box
        sx={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
        }}
      >
        <Button
          variant="outlined"
          startIcon={<NavigateBefore />}
          onClick={() => setActivePhase(4)}
        >
          이전 단계
        </Button>

        <Box sx={{ display: 'flex', gap: 2 }}>
          <Button
            variant="outlined"
            color="warning"
            startIcon={<RestartAlt />}
            onClick={handleReset}
          >
            새 프로젝트 시작
          </Button>

          {isComplete && (
            <Button
              variant="contained"
              color="success"
              startIcon={<CheckCircle />}
            >
              작업 완료
            </Button>
          )}
        </Box>
      </Box>
    </Box>
  )
}

export default FinalProcessingPhase
