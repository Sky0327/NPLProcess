import React from 'react'
import {
  Box,
  Typography,
  Button,
  Grid,
  Card,
  CardContent,
  LinearProgress,
  Tabs,
  Tab,
  Chip,
  Dialog,
  DialogTitle,
  DialogContent,
  IconButton,
} from '@mui/material'
import {
  PlayArrow,
  NavigateNext,
  NavigateBefore,
  Download,
  Preview,
  Close,
  ListAlt,
  Assessment,
  Print,
} from '@mui/icons-material'
import useWorkflowStore from '../../../store/workflowStore'
import { PHASES } from '../../../data/workflowConfig'
import TaskCard from '../../Common/TaskCard'
import EvaluationReportPreview from '../../Reports/EvaluationReportPreview'

const ReportPhase = () => {
  const [activeTab, setActiveTab] = React.useState(0)
  const [previewOpen, setPreviewOpen] = React.useState(false)

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

  const phase = PHASES[4]
  const progress = calculatePhaseProgress(4)
  const isComplete = progress === 100

  const handleGenerateReport = async (task) => {
    if (!canExecuteTask(task.id)) {
      addLogEntry({
        type: 'warning',
        action: `${task.name} 생성 불가`,
        details: `필요한 데이터: ${task.deps.join(', ')}`,
        phase: 4,
      })
      return
    }

    setTaskStatus(task.id, { loading: true, error: null })

    addLogEntry({
      type: 'info',
      action: `${task.name} 생성 시작`,
      phase: 4,
    })

    try {
      // Simulate report generation
      await new Promise((resolve) => setTimeout(resolve, 1500))

      const reportData = {
        name: task.name,
        generatedAt: new Date().toISOString(),
        status: 'completed',
      }

      setTaskResult(task.id, reportData)
      setTaskStatus(task.id, { loading: false, error: null })

      addLogEntry({
        type: 'success',
        action: `${task.name} 생성 완료`,
        phase: 4,
      })
    } catch (error) {
      setTaskStatus(task.id, { loading: false, error: error.message })

      addLogEntry({
        type: 'error',
        action: `${task.name} 생성 오류`,
        details: error.message,
        phase: 4,
      })
    }
  }

  const handleGenerateAll = async () => {
    const executableTasks = phase.tasks.filter((t) => canExecuteTask(t.id))

    if (executableTasks.length === 0) {
      addLogEntry({
        type: 'warning',
        action: '실행 가능한 리포트 없음',
        details: '필요한 데이터를 먼저 준비해주세요.',
        phase: 4,
      })
      return
    }

    addLogEntry({
      type: 'info',
      action: '전체 리포트 생성 시작',
      details: `${executableTasks.length}개 리포트 생성`,
      phase: 4,
    })

    for (const task of executableTasks) {
      await handleGenerateReport(task)
    }

    addLogEntry({
      type: 'success',
      action: '전체 리포트 생성 완료',
      phase: 4,
    })
  }

  const handlePrint = () => {
    window.print()
  }

  const isAnyLoading = phase.tasks.some((t) => taskStatus[t.id]?.loading)
  const completedReports = phase.tasks.filter((t) => taskResults[t.id])

  return (
    <Box>
      {/* Header */}
      <Box
        sx={{
          mb: 3,
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'flex-start',
          flexWrap: 'wrap',
          gap: 2,
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
            처리된 데이터를 기반으로 리포트를 생성하고 미리볼 수 있습니다.
          </Typography>
        </Box>

        <Box sx={{ display: 'flex', gap: 1 }}>
          <Button
            variant="outlined"
            startIcon={<Preview />}
            onClick={() => setPreviewOpen(true)}
          >
            평가보고서 미리보기
          </Button>
          <Button
            variant="contained"
            startIcon={<PlayArrow />}
            onClick={handleGenerateAll}
            disabled={isAnyLoading}
            sx={{ bgcolor: phase.color, '&:hover': { bgcolor: phase.color } }}
          >
            전체 생성
          </Button>
        </Box>
      </Box>

      {/* Tabs */}
      <Card sx={{ mb: 2 }}>
        <Tabs
          value={activeTab}
          onChange={(e, v) => setActiveTab(v)}
          variant="fullWidth"
          sx={{
            borderBottom: 1,
            borderColor: 'divider',
            '& .MuiTab-root': {
              textTransform: 'none',
              fontWeight: 500,
              minHeight: 48,
            },
          }}
        >
          <Tab
            icon={<ListAlt fontSize="small" />}
            iconPosition="start"
            label={
              <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                리포트 목록
                <Chip
                  label={`${completedReports.length}/${phase.tasks.length}`}
                  size="small"
                  sx={{ height: 20, fontSize: '0.75rem' }}
                />
              </Box>
            }
          />
          <Tab
            icon={<Assessment fontSize="small" />}
            iconPosition="start"
            label="평가보고서 미리보기"
          />
        </Tabs>
      </Card>

      {/* Tab Content: Report Tasks */}
      {activeTab === 0 && (
        <>
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
                  리포트 생성 진행률
                </Typography>
                <Typography variant="body2" sx={{ color: 'text.secondary' }}>
                  {completedReports.length} / {phase.tasks.length}
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

          {/* Report Grid */}
          <Grid container spacing={2}>
            {phase.tasks.map((task) => (
              <Grid item xs={12} sm={6} md={4} key={task.id}>
                <TaskCard
                  id={task.id}
                  name={task.name}
                  result={taskResults[task.id]}
                  loading={taskStatus[task.id]?.loading}
                  error={taskStatus[task.id]?.error}
                  dependencies={task.deps}
                  dependenciesMet={canExecuteTask(task.id)}
                  onExecute={() => handleGenerateReport(task)}
                  compact
                />
              </Grid>
            ))}
          </Grid>

          {/* Download Section */}
          {completedReports.length > 0 && (
            <Card sx={{ mt: 3, bgcolor: 'success.50' }}>
              <CardContent>
                <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <Box>
                    <Typography variant="subtitle2" sx={{ fontWeight: 600 }}>
                      생성된 리포트
                    </Typography>
                    <Typography variant="body2" color="text.secondary">
                      {completedReports.length}개의 리포트가 생성되었습니다.
                    </Typography>
                  </Box>
                  <Button variant="outlined" startIcon={<Download />}>
                    전체 다운로드
                  </Button>
                </Box>
              </CardContent>
            </Card>
          )}
        </>
      )}

      {/* Tab Content: Report Preview */}
      {activeTab === 1 && (
        <Box sx={{ mt: 2 }}>
          <Box sx={{ display: 'flex', justifyContent: 'flex-end', mb: 2, gap: 1 }}>
            <Button
              variant="outlined"
              startIcon={<Print />}
              onClick={handlePrint}
            >
              인쇄
            </Button>
            <Button
              variant="outlined"
              startIcon={<Download />}
            >
              PDF 저장
            </Button>
          </Box>
          <EvaluationReportPreview />
        </Box>
      )}

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
          onClick={() => setActivePhase(3)}
        >
          이전 단계
        </Button>
        <Button
          variant="contained"
          endIcon={<NavigateNext />}
          onClick={() => {
            addLogEntry({
              type: 'info',
              action: 'Phase 4 완료',
              details: '최종 처리 단계로 이동',
              phase: 4,
            })
            setActivePhase(5)
          }}
          disabled={!isComplete}
        >
          다음 단계
        </Button>
      </Box>

      {/* Full Screen Preview Dialog */}
      <Dialog
        open={previewOpen}
        onClose={() => setPreviewOpen(false)}
        maxWidth="lg"
        fullWidth
        PaperProps={{
          sx: { minHeight: '90vh' }
        }}
      >
        <DialogTitle sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <Typography variant="h6">NPL 평가보고서 미리보기</Typography>
          <Box sx={{ display: 'flex', gap: 1 }}>
            <Button
              variant="outlined"
              size="small"
              startIcon={<Print />}
              onClick={handlePrint}
            >
              인쇄
            </Button>
            <Button
              variant="outlined"
              size="small"
              startIcon={<Download />}
            >
              PDF 저장
            </Button>
            <IconButton onClick={() => setPreviewOpen(false)} size="small">
              <Close />
            </IconButton>
          </Box>
        </DialogTitle>
        <DialogContent dividers>
          <EvaluationReportPreview />
        </DialogContent>
      </Dialog>
    </Box>
  )
}

export default ReportPhase
