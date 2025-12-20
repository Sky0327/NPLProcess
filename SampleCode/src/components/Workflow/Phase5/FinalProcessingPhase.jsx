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
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  List,
  ListItem,
  ListItemIcon,
  ListItemText,
  Chip,
} from '@mui/material'
import {
  NavigateBefore,
  Description,
  PictureAsPdf,
  FileDownload,
  RestartAlt,
  CheckCircle,
  TableChart,
  Folder,
} from '@mui/icons-material'
import * as XLSX from 'xlsx'
import useWorkflowStore from '../../../store/workflowStore'
import { PHASES } from '../../../data/workflowConfig'
import TaskCard from '../../Common/TaskCard'
import {
  calculateAllResults,
  formatNumber,
  formatPercent,
} from '../../../utils/nplCalculations'

const FinalProcessingPhase = () => {
  const [exportDialogOpen, setExportDialogOpen] = React.useState(false)
  const [exportedFile, setExportedFile] = React.useState(null)

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
    projectConfig,
    coverData,
    debtInfo,
    collateralInfo,
  } = useWorkflowStore()

  const phase = PHASES[5]
  const progress = calculatePhaseProgress(5)
  const overallProgress = calculateOverallProgress()
  const isComplete = progress === 100

  // Calculate results for export
  const calculatedResults = React.useMemo(() => {
    return calculateAllResults(collateralInfo, debtInfo, coverData)
  }, [collateralInfo, debtInfo, coverData])

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

  // Excel Export Function
  const handleExcelExport = async () => {
    setTaskStatus('xlsxExport', { loading: true, error: null })

    addLogEntry({
      type: 'info',
      action: 'XLSX 내보내기 시작',
      phase: 5,
    })

    try {
      // Create workbook
      const workbook = XLSX.utils.book_new()

      // Sheet 1: Cover Sheet
      const coverSheetData = [
        ['NPL 평가 보고서'],
        [],
        ['보고서명', projectConfig.reportName],
        ['프로젝트 ID', projectConfig.projectId],
        ['작성일', coverData.reportDate || new Date().toLocaleDateString('ko-KR')],
        [],
        ['기본 정보'],
        ['매각기관', coverData.sellingInstitution],
        ['업무구분', coverData.businessType],
        ['차주명', coverData.borrowerName],
        ['사업자구분', coverData.businessClassification],
        ['물건유형', coverData.propertyType],
        ['소재지', coverData.address],
        ['작성기관', coverData.preparingOrg],
        [],
        ['주요 가정'],
        ['기본할인율', `${(coverData.baseRate * 100).toFixed(2)}%`],
        ['매입할인율', `${(coverData.purchaseRate * 100).toFixed(2)}%`],
        ['현가할인율', `${(coverData.discountRate * 100).toFixed(2)}%`],
        ['관리비용률', `${(coverData.managementCostRate * 100).toFixed(2)}%`],
        ['할인기간', `${coverData.discountPeriod}일`],
      ]
      const coverSheet = XLSX.utils.aoa_to_sheet(coverSheetData)
      coverSheet['!cols'] = [{ wch: 15 }, { wch: 50 }]
      XLSX.utils.book_append_sheet(workbook, coverSheet, 'Cover')

      // Sheet 2: Debt Status
      const debtHeaders = [
        '순위', '금융기관', '계좌번호', '채권최고액', '대출잔액',
        '가지급금', '미수이자', '연체이자율(%)', '취급일', '만기일'
      ]
      const debtRows = debtInfo.map((debt, idx) => [
        debt.priority || `${idx + 1}순위`,
        debt.institutionName,
        debt.accountNumber,
        debt.maxDebtAmount,
        debt.outstandingBalance,
        debt.advancePayment,
        debt.accruedInterest,
        debt.delinquencyRate,
        debt.handlingDate,
        debt.maturityDate,
      ])
      const debtTotals = [
        '합계', '', '',
        debtInfo.reduce((sum, d) => sum + (d.maxDebtAmount || 0), 0),
        debtInfo.reduce((sum, d) => sum + (d.outstandingBalance || 0), 0),
        debtInfo.reduce((sum, d) => sum + (d.advancePayment || 0), 0),
        debtInfo.reduce((sum, d) => sum + (d.accruedInterest || 0), 0),
        '', '', ''
      ]
      const debtSheetData = [debtHeaders, ...debtRows, debtTotals]
      const debtSheet = XLSX.utils.aoa_to_sheet(debtSheetData)
      debtSheet['!cols'] = [
        { wch: 8 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 },
        { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }
      ]
      XLSX.utils.book_append_sheet(workbook, debtSheet, '채권현황')

      // Sheet 3: Collateral Information
      const colHeaders = [
        '물건', '물건유형', '소재지', '용도지역', '토지면적(㎡)',
        '건물면적(㎡)', '감정평가액', '자체감정가', '낙찰가율(%)'
      ]
      const colRows = collateralInfo.map((col, idx) => [
        `물건 ${idx + 1}`,
        col.propertyType,
        col.address,
        col.landUseZone,
        col.landArea,
        col.buildingArea,
        col.appraisedValue,
        col.internalAppraisal,
        col.auctionRate,
      ])
      const colTotals = [
        '합계', '', '', '',
        collateralInfo.reduce((sum, c) => sum + (c.landArea || 0), 0),
        collateralInfo.reduce((sum, c) => sum + (c.buildingArea || 0), 0),
        collateralInfo.reduce((sum, c) => sum + (c.appraisedValue || 0), 0),
        collateralInfo.reduce((sum, c) => sum + (c.internalAppraisal || 0), 0),
        ''
      ]
      const colSheetData = [colHeaders, ...colRows, colTotals]
      const colSheet = XLSX.utils.aoa_to_sheet(colSheetData)
      colSheet['!cols'] = [
        { wch: 8 }, { wch: 12 }, { wch: 40 }, { wch: 12 }, { wch: 12 },
        { wch: 12 }, { wch: 15 }, { wch: 15 }, { wch: 12 }
      ]
      XLSX.utils.book_append_sheet(workbook, colSheet, '담보물정보')

      // Sheet 4: Calculation Results
      const resultsData = [
        ['NPL 평가 결과'],
        [],
        ['항목', '금액', '비고'],
        ['총 담보평가액', calculatedResults.totalCollateralValue, '감정평가액 × 낙찰가율'],
        ['총 회수시점채권액', calculatedResults.totalRecoveryDebt, ''],
        [],
        ['NRV (순회수금)', calculatedResults.totalNRV, 'MIN(배당가능액, 근저당권설정액, 회수시점채권액) - 이전비용'],
        ['NPV (매각대금)', calculatedResults.totalNPV, '(NRV - 관리비용) × 할인계수'],
        ['FV (미래가)', calculatedResults.totalFV, 'NPV × (1 + 현가할인율 × 할인기간/365)'],
        [],
        ['총 수수료', calculatedResults.totalFees, 'FV - NPV + 관리비용'],
        ['정산차액', calculatedResults.totalSettlement, '확정배당금 - FV - 정산비용 - 관리비용'],
        [],
        ['OPB 대비 회수율', calculatedResults.totalOutstandingBalance > 0
          ? (calculatedResults.totalNPV / calculatedResults.totalOutstandingBalance * 100).toFixed(2) + '%'
          : '-', 'NPV / (대출잔액 + 가지급금)'],
        ['채권최고액 대비', calculatedResults.totalMaxDebtAmount > 0
          ? (calculatedResults.totalNPV / calculatedResults.totalMaxDebtAmount * 100).toFixed(2) + '%'
          : '-', 'NPV / 채권최고액'],
      ]
      const resultsSheet = XLSX.utils.aoa_to_sheet(resultsData)
      resultsSheet['!cols'] = [{ wch: 20 }, { wch: 20 }, { wch: 40 }]
      XLSX.utils.book_append_sheet(workbook, resultsSheet, '평가결과')

      // Sheet 5: Property-wise Breakdown (if multiple properties)
      if (calculatedResults.propertyResults && calculatedResults.propertyResults.length > 1) {
        const propHeaders = [
          '물건', '담보평가액', 'NRV', 'NPV', 'FV', '수수료', '회수율'
        ]
        const propRows = calculatedResults.propertyResults.map((prop, idx) => [
          `물건 ${idx + 1}`,
          prop.collateralValue,
          prop.nrv,
          prop.npv,
          prop.fv,
          prop.totalFee,
          prop.recoveryRate ? (prop.recoveryRate * 100).toFixed(2) + '%' : '-'
        ])
        const propSheetData = [propHeaders, ...propRows]
        const propSheet = XLSX.utils.aoa_to_sheet(propSheetData)
        propSheet['!cols'] = [
          { wch: 10 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 12 }, { wch: 10 }
        ]
        XLSX.utils.book_append_sheet(workbook, propSheet, '물건별상세')
      }

      // Generate filename
      const filename = `NPL_평가보고서_${projectConfig.reportName || 'report'}_${new Date().toISOString().slice(0, 10)}.xlsx`

      // Write file
      XLSX.writeFile(workbook, filename)

      setExportedFile(filename)
      setTaskResult('xlsxExport', {
        completedAt: new Date().toISOString(),
        status: 'completed',
        filename: filename,
      })
      setTaskStatus('xlsxExport', { loading: false, error: null })

      addLogEntry({
        type: 'success',
        action: 'XLSX 내보내기 완료',
        details: `파일명: ${filename}`,
        phase: 5,
      })

      setExportDialogOpen(true)
    } catch (error) {
      setTaskStatus('xlsxExport', { loading: false, error: error.message })

      addLogEntry({
        type: 'error',
        action: 'XLSX 내보내기 오류',
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
      customAction: handleExcelExport,
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
              onExecute={task.customAction || (() => handleTask(task.id, task.name))}
              compact
            />
          </Grid>
        ))}
      </Grid>

      {/* Export Summary */}
      {exportedFile && (
        <Card sx={{ mb: 3, bgcolor: 'success.50', borderColor: 'success.200', borderWidth: 1, borderStyle: 'solid' }}>
          <CardContent>
            <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
              <TableChart color="success" />
              <Box sx={{ flex: 1 }}>
                <Typography variant="subtitle2" sx={{ fontWeight: 600 }}>
                  Excel 파일 생성 완료
                </Typography>
                <Typography variant="body2" color="text.secondary">
                  {exportedFile}
                </Typography>
              </Box>
              <Chip label="다운로드 완료" color="success" size="small" />
            </Box>
          </CardContent>
        </Card>
      )}

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

      {/* Export Success Dialog */}
      <Dialog
        open={exportDialogOpen}
        onClose={() => setExportDialogOpen(false)}
        maxWidth="sm"
        fullWidth
      >
        <DialogTitle sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
          <CheckCircle color="success" />
          Excel 내보내기 완료
        </DialogTitle>
        <DialogContent>
          <Typography variant="body2" sx={{ mb: 2 }}>
            NPL 평가보고서가 Excel 파일로 저장되었습니다.
          </Typography>
          <Card variant="outlined" sx={{ bgcolor: 'grey.50' }}>
            <CardContent>
              <List dense>
                <ListItem>
                  <ListItemIcon>
                    <Folder fontSize="small" />
                  </ListItemIcon>
                  <ListItemText
                    primary="파일명"
                    secondary={exportedFile}
                  />
                </ListItem>
                <ListItem>
                  <ListItemIcon>
                    <TableChart fontSize="small" />
                  </ListItemIcon>
                  <ListItemText
                    primary="포함된 시트"
                    secondary="Cover, 채권현황, 담보물정보, 평가결과, 물건별상세"
                  />
                </ListItem>
              </List>
            </CardContent>
          </Card>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setExportDialogOpen(false)}>
            닫기
          </Button>
        </DialogActions>
      </Dialog>
    </Box>
  )
}

export default FinalProcessingPhase
