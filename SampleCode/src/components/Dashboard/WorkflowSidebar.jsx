import React from 'react'
import {
  Box,
  Typography,
  List,
  ListItemButton,
  ListItemIcon,
  ListItemText,
  Collapse,
  Chip,
  LinearProgress,
} from '@mui/material'
import {
  ExpandLess,
  ExpandMore,
  CheckCircle,
  RadioButtonUnchecked,
  PlayCircle,
  Error as ErrorIcon,
} from '@mui/icons-material'
import useWorkflowStore from '../../store/workflowStore'
import { PHASES } from '../../data/workflowConfig'

const WorkflowSidebar = ({ open, width }) => {
  const {
    activePhase,
    setActivePhase,
    taskResults,
    taskStatus,
    calculatePhaseProgress,
  } = useWorkflowStore()

  const [expandedPhases, setExpandedPhases] = React.useState({
    1: true,
    2: true,
    3: false,
    4: false,
    5: false,
  })

  const togglePhase = (phaseId) => {
    setExpandedPhases((prev) => ({
      ...prev,
      [phaseId]: !prev[phaseId],
    }))
  }

  const getTaskIcon = (taskId) => {
    const status = taskStatus[taskId]
    const result = taskResults[taskId]

    if (status?.loading) {
      return <PlayCircle sx={{ color: 'warning.main', fontSize: 18 }} />
    }
    if (status?.error) {
      return <ErrorIcon sx={{ color: 'error.main', fontSize: 18 }} />
    }
    if (result !== null && result !== undefined) {
      return <CheckCircle sx={{ color: 'success.main', fontSize: 18 }} />
    }
    return <RadioButtonUnchecked sx={{ color: 'grey.400', fontSize: 18 }} />
  }

  const getPhaseStatus = (phaseId) => {
    const progress = calculatePhaseProgress(phaseId)
    if (progress === 100) return 'completed'
    if (progress > 0) return 'in_progress'
    return 'pending'
  }

  const getPhaseIcon = (phaseId) => {
    const status = getPhaseStatus(phaseId)
    const phase = PHASES[phaseId]

    if (status === 'completed') {
      return <CheckCircle sx={{ color: phase.color, fontSize: 20 }} />
    }
    if (status === 'in_progress') {
      return <PlayCircle sx={{ color: phase.color, fontSize: 20 }} />
    }
    return <RadioButtonUnchecked sx={{ color: 'grey.400', fontSize: 20 }} />
  }

  if (!open) return null

  return (
    <Box
      sx={{
        width: width,
        height: '100vh',
        bgcolor: '#1a1a2e',
        color: '#fff',
        display: 'flex',
        flexDirection: 'column',
        position: 'fixed',
        left: 0,
        top: 0,
      }}
    >
      {/* Header */}
      <Box
        sx={{
          p: 2,
          borderBottom: '1px solid rgba(255,255,255,0.1)',
        }}
      >
        <Typography
          variant="subtitle1"
          sx={{
            fontWeight: 600,
            color: 'rgba(255,255,255,0.9)',
          }}
        >
          워크플로우
        </Typography>
        <Typography
          variant="caption"
          sx={{
            color: 'rgba(255,255,255,0.5)',
          }}
        >
          5단계 통합 프로세스
        </Typography>
      </Box>

      {/* Phase List */}
      <Box
        sx={{
          flexGrow: 1,
          overflow: 'auto',
          py: 1,
          '&::-webkit-scrollbar': {
            width: 6,
          },
          '&::-webkit-scrollbar-thumb': {
            bgcolor: 'rgba(255,255,255,0.2)',
            borderRadius: 3,
          },
        }}
      >
        <List disablePadding>
          {Object.values(PHASES).map((phase) => {
            const progress = calculatePhaseProgress(phase.id)
            const isActive = activePhase === phase.id
            const isExpanded = expandedPhases[phase.id]

            return (
              <React.Fragment key={phase.id}>
                {/* Phase Header */}
                <ListItemButton
                  onClick={() => {
                    setActivePhase(phase.id)
                    togglePhase(phase.id)
                  }}
                  selected={isActive}
                  sx={{
                    mx: 1,
                    borderRadius: 1,
                    mb: 0.5,
                    '&.Mui-selected': {
                      bgcolor: 'rgba(255,255,255,0.1)',
                      '&:hover': {
                        bgcolor: 'rgba(255,255,255,0.15)',
                      },
                    },
                    '&:hover': {
                      bgcolor: 'rgba(255,255,255,0.05)',
                    },
                  }}
                >
                  <ListItemIcon sx={{ minWidth: 32 }}>
                    {getPhaseIcon(phase.id)}
                  </ListItemIcon>
                  <ListItemText
                    primary={
                      <Typography
                        variant="body2"
                        sx={{
                          fontWeight: isActive ? 600 : 400,
                          color: 'rgba(255,255,255,0.9)',
                        }}
                      >
                        {phase.name}
                      </Typography>
                    }
                    secondary={
                      progress > 0 && (
                        <Box sx={{ mt: 0.5 }}>
                          <LinearProgress
                            variant="determinate"
                            value={progress}
                            sx={{
                              height: 3,
                              borderRadius: 1,
                              bgcolor: 'rgba(255,255,255,0.1)',
                              '& .MuiLinearProgress-bar': {
                                bgcolor: phase.color,
                              },
                            }}
                          />
                        </Box>
                      )
                    }
                  />
                  {phase.tasks.length > 0 && (
                    <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                      <Chip
                        label={`${progress}%`}
                        size="small"
                        sx={{
                          height: 18,
                          fontSize: '0.65rem',
                          bgcolor: 'rgba(255,255,255,0.1)',
                          color: 'rgba(255,255,255,0.7)',
                        }}
                      />
                      {isExpanded ? (
                        <ExpandLess
                          sx={{ color: 'rgba(255,255,255,0.5)', fontSize: 18 }}
                        />
                      ) : (
                        <ExpandMore
                          sx={{ color: 'rgba(255,255,255,0.5)', fontSize: 18 }}
                        />
                      )}
                    </Box>
                  )}
                </ListItemButton>

                {/* Task List */}
                <Collapse in={isExpanded} timeout="auto" unmountOnExit>
                  <List disablePadding sx={{ pl: 2 }}>
                    {phase.tasks.map((task) => (
                      <ListItemButton
                        key={task.id}
                        sx={{
                          py: 0.5,
                          mx: 1,
                          borderRadius: 1,
                          '&:hover': {
                            bgcolor: 'rgba(255,255,255,0.05)',
                          },
                        }}
                      >
                        <ListItemIcon sx={{ minWidth: 28 }}>
                          {getTaskIcon(task.id)}
                        </ListItemIcon>
                        <ListItemText
                          primary={
                            <Typography
                              variant="caption"
                              sx={{
                                color: 'rgba(255,255,255,0.7)',
                                fontSize: '0.75rem',
                              }}
                            >
                              {task.name}
                            </Typography>
                          }
                        />
                      </ListItemButton>
                    ))}
                  </List>
                </Collapse>
              </React.Fragment>
            )
          })}
        </List>
      </Box>

      {/* Footer */}
      <Box
        sx={{
          p: 2,
          borderTop: '1px solid rgba(255,255,255,0.1)',
        }}
      >
        <Typography
          variant="caption"
          sx={{
            color: 'rgba(255,255,255,0.4)',
            fontSize: '0.65rem',
          }}
        >
          Samil NPL 평가 시스템 v1.0
        </Typography>
      </Box>
    </Box>
  )
}

export default WorkflowSidebar
