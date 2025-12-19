import React from 'react'
import { Box, Typography, Tooltip, useTheme } from '@mui/material'
import useWorkflowStore from '../../store/workflowStore'
import { PHASES } from '../../data/workflowConfig'

const GlobalProgressBar = () => {
  const theme = useTheme()
  const { activePhase, calculatePhaseProgress, calculateOverallProgress } =
    useWorkflowStore()

  const overallProgress = calculateOverallProgress()

  const phaseSegments = Object.values(PHASES).map((phase) => {
    const progress = calculatePhaseProgress(phase.id)
    const isActive = activePhase === phase.id
    const isCompleted = progress === 100

    return {
      ...phase,
      progress,
      isActive,
      isCompleted,
    }
  })

  return (
    <Box
      sx={{
        px: 2,
        py: 1.5,
        bgcolor: '#f8f9fa',
        borderBottom: '1px solid',
        borderColor: 'divider',
      }}
    >
      {/* Progress Bar Container */}
      <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
        <Box sx={{ flexGrow: 1 }}>
          <Box
            sx={{
              display: 'flex',
              height: 28,
              borderRadius: 2,
              overflow: 'hidden',
              bgcolor: '#e9ecef',
            }}
          >
            {phaseSegments.map((phase, index) => (
              <Tooltip
                key={phase.id}
                title={`${phase.name}: ${phase.progress}%`}
                arrow
              >
                <Box
                  sx={{
                    width: `${phase.weight}%`,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    position: 'relative',
                    bgcolor: phase.isCompleted
                      ? phase.color
                      : phase.isActive
                      ? `${phase.color}40`
                      : '#e9ecef',
                    borderRight:
                      index < 4 ? '2px solid rgba(255,255,255,0.5)' : 'none',
                    transition: 'background-color 0.3s ease',
                    cursor: 'default',
                    overflow: 'hidden',
                  }}
                >
                  {/* Progress fill within segment */}
                  {!phase.isCompleted && phase.progress > 0 && (
                    <Box
                      sx={{
                        position: 'absolute',
                        left: 0,
                        top: 0,
                        bottom: 0,
                        width: `${phase.progress}%`,
                        bgcolor: phase.color,
                        transition: 'width 0.3s ease',
                      }}
                    />
                  )}
                  <Typography
                    variant="caption"
                    sx={{
                      color:
                        phase.isCompleted || phase.progress > 50
                          ? '#fff'
                          : '#666',
                      fontWeight: phase.isActive ? 600 : 400,
                      fontSize: '0.7rem',
                      position: 'relative',
                      zIndex: 1,
                      textShadow:
                        phase.isCompleted || phase.progress > 50
                          ? '0 1px 2px rgba(0,0,0,0.2)'
                          : 'none',
                    }}
                  >
                    P{phase.id}
                  </Typography>
                </Box>
              </Tooltip>
            ))}
          </Box>
        </Box>

        {/* Overall Progress */}
        <Box
          sx={{
            minWidth: 80,
            textAlign: 'right',
          }}
        >
          <Typography
            variant="subtitle2"
            sx={{
              fontWeight: 600,
              color: 'text.primary',
              fontSize: '0.85rem',
            }}
          >
            {overallProgress}%
          </Typography>
          <Typography
            variant="caption"
            sx={{
              color: 'text.secondary',
              fontSize: '0.65rem',
            }}
          >
            전체 진행률
          </Typography>
        </Box>
      </Box>

      {/* Phase Labels */}
      <Box
        sx={{
          display: 'flex',
          mt: 0.5,
        }}
      >
        {phaseSegments.map((phase) => (
          <Box
            key={phase.id}
            sx={{
              width: `${phase.weight}%`,
              textAlign: 'center',
            }}
          >
            <Typography
              variant="caption"
              sx={{
                fontSize: '0.65rem',
                color: phase.isActive ? 'primary.main' : 'text.secondary',
                fontWeight: phase.isActive ? 600 : 400,
              }}
            >
              {phase.name}
            </Typography>
          </Box>
        ))}
      </Box>
    </Box>
  )
}

export default GlobalProgressBar
