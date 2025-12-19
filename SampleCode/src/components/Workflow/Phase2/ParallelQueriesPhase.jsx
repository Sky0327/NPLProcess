import React from 'react'
import {
  Box,
  Typography,
  Button,
  Grid,
  Card,
  CardContent,
  LinearProgress,
} from '@mui/material'
import { PlayArrow, NavigateNext, NavigateBefore } from '@mui/icons-material'
import useWorkflowStore from '../../../store/workflowStore'
import { PHASES } from '../../../data/workflowConfig'
import { apiService } from '../../../services/apiService'
import TaskCard from '../../Common/TaskCard'

const ParallelQueriesPhase = () => {
  const {
    taskResults,
    taskStatus,
    setTaskResult,
    setTaskStatus,
    setActivePhase,
    addLogEntry,
    calculatePhaseProgress,
    projectConfig,
  } = useWorkflowStore()

  const phase = PHASES[2]
  const progress = calculatePhaseProgress(2)
  const isComplete = progress === 100

  const handleExecuteTask = async (task) => {
    setTaskStatus(task.id, { loading: true, error: null })

    addLogEntry({
      type: 'info',
      action: `${task.name} 시작`,
      phase: 2,
    })

    try {
      const result = await apiService[task.api](projectConfig.inputData || [])

      if (result.success) {
        setTaskResult(task.id, result.data)
        setTaskStatus(task.id, { loading: false, error: null })

        addLogEntry({
          type: 'success',
          action: `${task.name} 완료`,
          details: `${result.data?.length || 0}건 조회`,
          phase: 2,
        })
      } else {
        throw new Error(result.error || '조회 실패')
      }
    } catch (error) {
      setTaskStatus(task.id, { loading: false, error: error.message })

      addLogEntry({
        type: 'error',
        action: `${task.name} 오류`,
        details: error.message,
        phase: 2,
      })
    }
  }

  const handleExecuteAll = async () => {
    addLogEntry({
      type: 'info',
      action: '전체 조회 시작',
      details: `${phase.tasks.length}개 프로세스 병렬 실행`,
      phase: 2,
    })

    const promises = phase.tasks.map((task) => handleExecuteTask(task))
    await Promise.all(promises)

    addLogEntry({
      type: 'success',
      action: '전체 조회 완료',
      phase: 2,
    })
  }

  const isAnyLoading = phase.tasks.some((t) => taskStatus[t.id]?.loading)

  return (
    <Box>
      {/* Header */}
      <Box
        sx={{
          mb: 3,
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'flex-start',
        }}
      >
        <Box>
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
            7개 데이터 소스에서 병렬로 데이터를 조회합니다.
          </Typography>
        </Box>

        <Button
          variant="contained"
          startIcon={<PlayArrow />}
          onClick={handleExecuteAll}
          disabled={isAnyLoading}
          sx={{ bgcolor: phase.color, '&:hover': { bgcolor: phase.color } }}
        >
          전체 실행
        </Button>
      </Box>

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
          <Grid item xs={12} sm={6} md={4} lg={3} key={task.id}>
            <TaskCard
              id={task.id}
              name={task.name}
              result={taskResults[task.id]}
              loading={taskStatus[task.id]?.loading}
              error={taskStatus[task.id]?.error}
              onExecute={() => handleExecuteTask(task)}
              compact
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
          onClick={() => setActivePhase(1)}
        >
          이전 단계
        </Button>
        <Button
          variant="contained"
          endIcon={<NavigateNext />}
          onClick={() => {
            addLogEntry({
              type: 'info',
              action: 'Phase 2 완료',
              details: '중간 처리 단계로 이동',
              phase: 2,
            })
            setActivePhase(3)
          }}
          disabled={!isComplete}
        >
          다음 단계
        </Button>
      </Box>
    </Box>
  )
}

export default ParallelQueriesPhase
