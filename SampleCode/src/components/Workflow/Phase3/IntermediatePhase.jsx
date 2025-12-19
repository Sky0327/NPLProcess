import React from 'react'
import {
  Box,
  Typography,
  Button,
  Grid,
  Card,
  CardContent,
  LinearProgress,
  Alert,
} from '@mui/material'
import { NavigateNext, NavigateBefore } from '@mui/icons-material'
import useWorkflowStore from '../../../store/workflowStore'
import { PHASES } from '../../../data/workflowConfig'
import { apiService } from '../../../services/apiService'
import TaskCard from '../../Common/TaskCard'

const IntermediatePhase = () => {
  const {
    taskResults,
    taskStatus,
    setTaskResult,
    setTaskStatus,
    setActivePhase,
    addLogEntry,
    calculatePhaseProgress,
    canExecuteTask,
  } = useWorkflowStore()

  const phase = PHASES[3]
  const progress = calculatePhaseProgress(3)
  const isComplete = progress === 100

  const handleExecuteTask = async (task) => {
    if (!canExecuteTask(task.id)) {
      addLogEntry({
        type: 'warning',
        action: `${task.name} 실행 불가`,
        details: `필요한 데이터: ${task.deps.join(', ')}`,
        phase: 3,
      })
      return
    }

    setTaskStatus(task.id, { loading: true, error: null })

    addLogEntry({
      type: 'info',
      action: `${task.name} 시작`,
      phase: 3,
    })

    try {
      // Gather dependency data
      const depData = task.deps.reduce((acc, dep) => {
        acc[dep] = taskResults[dep]
        return acc
      }, {})

      const result = await apiService[task.api](depData)

      if (result.success) {
        setTaskResult(task.id, result.data)
        setTaskStatus(task.id, { loading: false, error: null })

        addLogEntry({
          type: 'success',
          action: `${task.name} 완료`,
          details: `${result.data?.length || 0}건 처리`,
          phase: 3,
        })
      } else {
        throw new Error(result.error || '처리 실패')
      }
    } catch (error) {
      setTaskStatus(task.id, { loading: false, error: error.message })

      addLogEntry({
        type: 'error',
        action: `${task.name} 오류`,
        details: error.message,
        phase: 3,
      })
    }
  }

  const isAnyLoading = phase.tasks.some((t) => taskStatus[t.id]?.loading)

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
          Phase 2에서 수집한 데이터를 기반으로 추가 처리를 수행합니다.
        </Typography>
      </Box>

      {/* Info Alert */}
      <Alert severity="info" sx={{ mb: 3 }}>
        각 프로세스는 의존성이 있어 순차적으로 실행해야 합니다. 잠금 아이콘이
        표시된 항목은 필요한 데이터가 아직 준비되지 않았습니다.
      </Alert>

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
              진행률
            </Typography>
            <Typography variant="body2" sx={{ color: 'text.secondary' }}>
              {phase.tasks.filter((t) => taskResults[t.id]).length} /{' '}
              {phase.tasks.length}
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

      {/* Task Grid */}
      <Grid container spacing={2}>
        {phase.tasks.map((task) => (
          <Grid item xs={12} sm={6} key={task.id}>
            <TaskCard
              id={task.id}
              name={task.name}
              result={taskResults[task.id]}
              loading={taskStatus[task.id]?.loading}
              error={taskStatus[task.id]?.error}
              dependencies={task.deps}
              dependenciesMet={canExecuteTask(task.id)}
              onExecute={() => handleExecuteTask(task)}
            />
          </Grid>
        ))}
      </Grid>

      {/* Navigation */}
      <Box
        sx={{
          mt: 4,
          display: 'flex',
          justifyContent: 'space-between',
        }}
      >
        <Button
          variant="outlined"
          startIcon={<NavigateBefore />}
          onClick={() => setActivePhase(2)}
        >
          이전 단계
        </Button>
        <Button
          variant="contained"
          endIcon={<NavigateNext />}
          onClick={() => {
            addLogEntry({
              type: 'info',
              action: 'Phase 3 완료',
              details: '리포트 생성 단계로 이동',
              phase: 3,
            })
            setActivePhase(4)
          }}
          disabled={!isComplete}
        >
          다음 단계
        </Button>
      </Box>
    </Box>
  )
}

export default IntermediatePhase
