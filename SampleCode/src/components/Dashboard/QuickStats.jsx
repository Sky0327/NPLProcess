import React from 'react'
import { Box, Card, CardContent, Typography, Grid } from '@mui/material'
import {
  Search,
  Settings,
  Description,
  Timer,
} from '@mui/icons-material'
import useWorkflowStore from '../../store/workflowStore'
import { PHASES } from '../../data/workflowConfig'

const StatCard = ({ icon: Icon, label, value, total, color }) => {
  const percentage = total ? Math.round((value / total) * 100) : 0

  return (
    <Card
      variant="compact"
      sx={{
        height: '100%',
        borderLeft: '3px solid',
        borderLeftColor: color,
      }}
    >
      <CardContent sx={{ p: 1.5, '&:last-child': { pb: 1.5 } }}>
        <Box sx={{ display: 'flex', alignItems: 'flex-start', gap: 1 }}>
          <Icon sx={{ color: color, fontSize: 20, mt: 0.25 }} />
          <Box>
            <Typography
              variant="caption"
              sx={{ color: 'text.secondary', display: 'block' }}
            >
              {label}
            </Typography>
            <Box sx={{ display: 'flex', alignItems: 'baseline', gap: 0.5 }}>
              <Typography variant="h6" sx={{ fontWeight: 700, lineHeight: 1 }}>
                {value}
              </Typography>
              {total && (
                <Typography variant="caption" sx={{ color: 'text.secondary' }}>
                  / {total}
                </Typography>
              )}
            </Box>
            {total && (
              <Box
                sx={{
                  mt: 0.5,
                  height: 3,
                  bgcolor: 'grey.200',
                  borderRadius: 1,
                  overflow: 'hidden',
                }}
              >
                <Box
                  sx={{
                    width: `${percentage}%`,
                    height: '100%',
                    bgcolor: color,
                    transition: 'width 0.3s ease',
                  }}
                />
              </Box>
            )}
          </Box>
        </Box>
      </CardContent>
    </Card>
  )
}

const QuickStats = () => {
  const { taskResults, activityLog } = useWorkflowStore()

  // Calculate stats
  const phase2Tasks = PHASES[2].tasks
  const phase3Tasks = PHASES[3].tasks
  const phase4Tasks = PHASES[4].tasks
  const phase5Tasks = PHASES[5].tasks

  const completedQueries = phase2Tasks.filter(
    (t) => taskResults[t.id] !== null
  ).length

  const completedProcessing = phase3Tasks.filter(
    (t) => taskResults[t.id] !== null
  ).length

  const completedReports = [...phase4Tasks, ...phase5Tasks].filter(
    (t) => taskResults[t.id] !== null
  ).length

  // Calculate elapsed time from first log entry
  const getElapsedTime = () => {
    if (activityLog.length === 0) return '00:00:00'

    const firstEntry = activityLog[activityLog.length - 1]
    const startTime = new Date(firstEntry.timestamp)
    const now = new Date()
    const diff = now - startTime

    const hours = Math.floor(diff / (1000 * 60 * 60))
    const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60))
    const seconds = Math.floor((diff % (1000 * 60)) / 1000)

    return `${hours.toString().padStart(2, '0')}:${minutes
      .toString()
      .padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`
  }

  return (
    <Box sx={{ mb: 3 }}>
      <Typography
        variant="subtitle2"
        sx={{ mb: 1.5, color: 'text.secondary', fontWeight: 500 }}
      >
        진행 현황
      </Typography>
      <Grid container spacing={2}>
        <Grid item xs={6} sm={3}>
          <StatCard
            icon={Search}
            label="데이터 조회"
            value={completedQueries}
            total={phase2Tasks.length}
            color="#2196f3"
          />
        </Grid>
        <Grid item xs={6} sm={3}>
          <StatCard
            icon={Settings}
            label="중간 처리"
            value={completedProcessing}
            total={phase3Tasks.length}
            color="#9c27b0"
          />
        </Grid>
        <Grid item xs={6} sm={3}>
          <StatCard
            icon={Description}
            label="리포트"
            value={completedReports}
            total={phase4Tasks.length + phase5Tasks.length}
            color="#ff9800"
          />
        </Grid>
        <Grid item xs={6} sm={3}>
          <StatCard
            icon={Timer}
            label="소요 시간"
            value={getElapsedTime()}
            color="#4caf50"
          />
        </Grid>
      </Grid>
    </Box>
  )
}

export default QuickStats
