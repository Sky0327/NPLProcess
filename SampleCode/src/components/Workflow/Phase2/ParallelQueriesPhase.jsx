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
  Dialog,
  DialogTitle,
  DialogContent,
  DialogContentText,
  DialogActions,
  Chip,
} from '@mui/material'
import {
  PlayArrow,
  NavigateNext,
  NavigateBefore,
  AutoFixHigh,
  Check,
} from '@mui/icons-material'
import useWorkflowStore from '../../../store/workflowStore'
import { PHASES } from '../../../data/workflowConfig'
import { apiService } from '../../../services/apiService'
import TaskCard from '../../Common/TaskCard'

const ParallelQueriesPhase = () => {
  const [autoFillDialogOpen, setAutoFillDialogOpen] = React.useState(false)
  const [autoFillCompleted, setAutoFillCompleted] = React.useState(false)

  const {
    taskResults,
    taskStatus,
    setTaskResult,
    setTaskStatus,
    setActivePhase,
    addLogEntry,
    calculatePhaseProgress,
    projectConfig,
    coverData,
    setCoverData,
    setDebtInfo,
    setCollateralInfo,
    debtInfo,
    collateralInfo,
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

  // Auto-fill Phase 1 data from query results
  const handleAutoFill = () => {
    setAutoFillDialogOpen(true)
  }

  const performAutoFill = () => {
    const generateId = () => `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`

    // Auto-fill from registry data
    const registryData = taskResults.registry
    const auctionData = taskResults.auction
    const landPriceData = taskResults.landPrice

    // Create collateral info from registry/land price data
    if (registryData && Array.isArray(registryData) && registryData.length > 0) {
      const newCollateralInfo = registryData.map((item, index) => ({
        id: generateId(),
        propertyIndex: index + 1,
        propertyType: item.propertyType || coverData.propertyType || '',
        address: item.address || coverData.address || '',
        landUseZone: item.landUseZone || '',
        landArea: item.landArea || 0,
        landPrice: landPriceData?.[index]?.price || item.landPrice || 0,
        buildingArea: item.buildingArea || 0,
        buildingScale: item.buildingScale || '',
        buildingStructure: item.buildingStructure || '',
        appraisedValue: item.appraisedValue || 0,
        internalAppraisal: item.internalAppraisal || item.appraisedValue || 0,
        appraiser: item.appraiser || '',
        auctionRate: auctionData?.[index]?.auctionRate || 80,
      }))

      if (newCollateralInfo.length > 0) {
        setCollateralInfo(newCollateralInfo)
      }

      // Update cover data with first property address if empty
      if (!coverData.address && newCollateralInfo[0]?.address) {
        setCoverData({ address: newCollateralInfo[0].address })
      }
    }

    // Create debt info from registry data (mortgage/debt information)
    if (registryData && Array.isArray(registryData) && registryData.length > 0) {
      const debtItems = registryData.flatMap((item) => {
        if (item.debts && Array.isArray(item.debts)) {
          return item.debts.map((debt, idx) => ({
            id: generateId(),
            priority: debt.priority || `${idx + 1}순위`,
            institutionName: debt.institutionName || debt.creditor || '',
            accountNumber: debt.accountNumber || '',
            maxDebtAmount: debt.maxDebtAmount || debt.mortgageAmount || 0,
            outstandingBalance: debt.outstandingBalance || debt.balance || 0,
            advancePayment: debt.advancePayment || 0,
            accruedInterest: debt.accruedInterest || 0,
            handlingDate: debt.handlingDate || debt.registrationDate || null,
            maturityDate: debt.maturityDate || null,
            defaultDate: debt.defaultDate || null,
            mortgageDate: debt.mortgageDate || debt.registrationDate || null,
            delinquencyRate: debt.delinquencyRate || 12,
          }))
        }
        // Single debt from property
        if (item.maxDebtAmount || item.mortgageAmount) {
          return [{
            id: generateId(),
            priority: '1순위',
            institutionName: item.creditor || coverData.sellingInstitution || '',
            accountNumber: item.accountNumber || '',
            maxDebtAmount: item.maxDebtAmount || item.mortgageAmount || 0,
            outstandingBalance: item.outstandingBalance || item.balance || 0,
            advancePayment: item.advancePayment || 0,
            accruedInterest: item.accruedInterest || 0,
            handlingDate: item.handlingDate || item.registrationDate || null,
            maturityDate: item.maturityDate || null,
            defaultDate: item.defaultDate || null,
            mortgageDate: item.mortgageDate || item.registrationDate || null,
            delinquencyRate: item.delinquencyRate || 12,
          }]
        }
        return []
      })

      if (debtItems.length > 0) {
        setDebtInfo(debtItems)
      }
    }

    // Log auto-fill action
    addLogEntry({
      type: 'success',
      action: '기본값 자동 입력 완료',
      details: `조회 결과에서 담보/채권 정보 추출`,
      phase: 2,
    })

    setAutoFillCompleted(true)
    setAutoFillDialogOpen(false)
  }

  const isAnyLoading = phase.tasks.some((t) => taskStatus[t.id]?.loading)
  const hasQueryResults = Object.keys(taskResults).some(
    (key) => taskResults[key] !== null && ['registry', 'landPrice', 'auction'].includes(key)
  )

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
            7개 데이터 소스에서 병렬로 데이터를 조회합니다.
          </Typography>
        </Box>

        <Box sx={{ display: 'flex', gap: 1 }}>
          <Button
            variant="outlined"
            startIcon={autoFillCompleted ? <Check /> : <AutoFixHigh />}
            onClick={handleAutoFill}
            disabled={!hasQueryResults || isAnyLoading}
            color={autoFillCompleted ? 'success' : 'primary'}
          >
            {autoFillCompleted ? '자동입력 완료' : '기본값 자동입력'}
          </Button>
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
      </Box>

      {/* Auto-fill Status */}
      {autoFillCompleted && (
        <Alert severity="success" sx={{ mb: 2 }}>
          기본값이 자동으로 입력되었습니다.{' '}
          <Chip
            label={`담보물 ${collateralInfo.length}건`}
            size="small"
            sx={{ ml: 1 }}
          />
          <Chip
            label={`채권 ${debtInfo.length}건`}
            size="small"
            sx={{ ml: 1 }}
          />
          <Button
            size="small"
            onClick={() => setActivePhase(1)}
            sx={{ ml: 2 }}
          >
            기본값 설정에서 확인
          </Button>
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

      {/* Auto-fill Confirmation Dialog */}
      <Dialog
        open={autoFillDialogOpen}
        onClose={() => setAutoFillDialogOpen(false)}
      >
        <DialogTitle>기본값 자동입력</DialogTitle>
        <DialogContent>
          <DialogContentText>
            조회 결과를 기반으로 채권현황과 담보물 정보를 자동으로 입력합니다.
          </DialogContentText>
          <Box sx={{ mt: 2 }}>
            <Typography variant="body2" color="text.secondary">
              자동입력 항목:
            </Typography>
            <ul style={{ margin: '8px 0', paddingLeft: 20 }}>
              <li>담보물 소재지, 면적, 감정평가액</li>
              <li>채권최고액, 대출잔액, 금융기관명</li>
              <li>낙찰가율 (경매정보 기준)</li>
              <li>개별공시지가 (공시지가 조회 기준)</li>
            </ul>
            <Alert severity="warning" sx={{ mt: 1 }}>
              기존에 입력된 데이터가 있으면 덮어쓰게 됩니다.
            </Alert>
          </Box>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setAutoFillDialogOpen(false)}>취소</Button>
          <Button onClick={performAutoFill} variant="contained" autoFocus>
            자동입력 실행
          </Button>
        </DialogActions>
      </Dialog>
    </Box>
  )
}

export default ParallelQueriesPhase
