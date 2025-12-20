import React from 'react'
import {
  Box,
  Card,
  CardContent,
  Typography,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  Grid,
  Divider,
  Chip,
} from '@mui/material'
import {
  Business,
  Person,
  Home,
  AccountBalance,
  Calculate,
  Assessment,
} from '@mui/icons-material'
import useWorkflowStore from '../../store/workflowStore'
import {
  calculateAllResults,
  formatNumber,
  formatPercent,
  formatKRW,
} from '../../utils/nplCalculations'

// Section Header Component
const SectionHeader = ({ icon: Icon, title, color = 'primary' }) => (
  <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 2 }}>
    <Icon color={color} />
    <Typography variant="h6" sx={{ fontWeight: 600 }}>
      {title}
    </Typography>
  </Box>
)

// Key-Value Display Component
const InfoRow = ({ label, value, bold = false }) => (
  <Box sx={{ display: 'flex', py: 0.5 }}>
    <Typography
      variant="body2"
      sx={{ width: 140, color: 'text.secondary', flexShrink: 0 }}
    >
      {label}
    </Typography>
    <Typography
      variant="body2"
      sx={{ fontWeight: bold ? 600 : 400, flex: 1 }}
    >
      {value || '-'}
    </Typography>
  </Box>
)

const EvaluationReportPreview = () => {
  const {
    projectConfig,
    coverData,
    debtInfo,
    collateralInfo,
  } = useWorkflowStore()

  // Calculate all results
  const results = React.useMemo(() => {
    return calculateAllResults(collateralInfo, debtInfo, coverData)
  }, [collateralInfo, debtInfo, coverData])

  // Debt totals
  const debtTotals = React.useMemo(() => ({
    maxDebtAmount: debtInfo.reduce((sum, d) => sum + (d.maxDebtAmount || 0), 0),
    outstandingBalance: debtInfo.reduce((sum, d) => sum + (d.outstandingBalance || 0), 0),
    advancePayment: debtInfo.reduce((sum, d) => sum + (d.advancePayment || 0), 0),
    accruedInterest: debtInfo.reduce((sum, d) => sum + (d.accruedInterest || 0), 0),
  }), [debtInfo])

  // Collateral totals
  const collateralTotals = React.useMemo(() => ({
    appraisedValue: collateralInfo.reduce((sum, c) => sum + (c.appraisedValue || 0), 0),
    internalAppraisal: collateralInfo.reduce((sum, c) => sum + (c.internalAppraisal || 0), 0),
    buildingArea: collateralInfo.reduce((sum, c) => sum + (c.buildingArea || 0), 0),
    landArea: collateralInfo.reduce((sum, c) => sum + (c.landArea || 0), 0),
  }), [collateralInfo])

  return (
    <Box sx={{ maxWidth: 1200, mx: 'auto' }}>
      {/* Report Header */}
      <Card sx={{ mb: 3, bgcolor: 'primary.900', color: 'white' }}>
        <CardContent>
          <Box sx={{ textAlign: 'center', py: 2 }}>
            <Typography variant="overline" sx={{ opacity: 0.8 }}>
              NPL 평가 보고서
            </Typography>
            <Typography variant="h4" sx={{ fontWeight: 700, mt: 1 }}>
              {projectConfig.reportName || '보고서명 미입력'}
            </Typography>
            <Typography variant="body2" sx={{ mt: 1, opacity: 0.8 }}>
              프로젝트 ID: {projectConfig.projectId || '-'} | 작성일: {coverData.reportDate || new Date().toLocaleDateString('ko-KR')}
            </Typography>
          </Box>
        </CardContent>
      </Card>

      <Grid container spacing={3}>
        {/* Cover Sheet Section */}
        <Grid item xs={12} md={6}>
          <Card sx={{ height: '100%' }}>
            <CardContent>
              <SectionHeader icon={Business} title="기본 정보" />
              <InfoRow label="매각기관" value={coverData.sellingInstitution} bold />
              <InfoRow label="업무구분" value={coverData.businessType} />
              <InfoRow label="차주명" value={coverData.borrowerName} bold />
              <InfoRow label="사업자구분" value={coverData.businessClassification} />
              <InfoRow label="물건유형" value={coverData.propertyType} />
              <InfoRow label="소재지" value={coverData.address} />
              <InfoRow label="작성기관" value={coverData.preparingOrg} />
            </CardContent>
          </Card>
        </Grid>

        {/* Key Assumptions Section */}
        <Grid item xs={12} md={6}>
          <Card sx={{ height: '100%' }}>
            <CardContent>
              <SectionHeader icon={Calculate} title="주요 가정" />
              <InfoRow
                label="기본할인율"
                value={`${(coverData.baseRate * 100).toFixed(2)}%`}
                bold
              />
              <InfoRow
                label="매입할인율"
                value={`${(coverData.purchaseRate * 100).toFixed(2)}%`}
              />
              <InfoRow
                label="현가할인율"
                value={`${(coverData.discountRate * 100).toFixed(2)}%`}
                bold
              />
              <InfoRow
                label="관리비용률"
                value={`${(coverData.managementCostRate * 100).toFixed(2)}%`}
              />
              <InfoRow
                label="할인기간"
                value={`${coverData.discountPeriod}일`}
                bold
              />
            </CardContent>
          </Card>
        </Grid>

        {/* Debt Status Section */}
        <Grid item xs={12}>
          <Card>
            <CardContent>
              <SectionHeader icon={AccountBalance} title="채권현황" />
              <TableContainer component={Paper} variant="outlined">
                <Table size="small">
                  <TableHead>
                    <TableRow sx={{ bgcolor: 'grey.100' }}>
                      <TableCell sx={{ fontWeight: 600 }}>순위</TableCell>
                      <TableCell sx={{ fontWeight: 600 }}>금융기관</TableCell>
                      <TableCell sx={{ fontWeight: 600 }}>계좌번호</TableCell>
                      <TableCell sx={{ fontWeight: 600 }} align="right">채권최고액</TableCell>
                      <TableCell sx={{ fontWeight: 600 }} align="right">대출잔액</TableCell>
                      <TableCell sx={{ fontWeight: 600 }} align="right">가지급금</TableCell>
                      <TableCell sx={{ fontWeight: 600 }} align="right">미수이자</TableCell>
                      <TableCell sx={{ fontWeight: 600 }} align="right">연체이자율</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {debtInfo.map((debt, index) => (
                      <TableRow key={debt.id}>
                        <TableCell>{debt.priority || `${index + 1}순위`}</TableCell>
                        <TableCell>{debt.institutionName || '-'}</TableCell>
                        <TableCell>{debt.accountNumber || '-'}</TableCell>
                        <TableCell align="right">{formatNumber(debt.maxDebtAmount)}</TableCell>
                        <TableCell align="right">{formatNumber(debt.outstandingBalance)}</TableCell>
                        <TableCell align="right">{formatNumber(debt.advancePayment)}</TableCell>
                        <TableCell align="right">{formatNumber(debt.accruedInterest)}</TableCell>
                        <TableCell align="right">{debt.delinquencyRate}%</TableCell>
                      </TableRow>
                    ))}
                    <TableRow sx={{ bgcolor: 'primary.50' }}>
                      <TableCell colSpan={3} sx={{ fontWeight: 600 }}>합계</TableCell>
                      <TableCell align="right" sx={{ fontWeight: 600 }}>
                        {formatNumber(debtTotals.maxDebtAmount)}
                      </TableCell>
                      <TableCell align="right" sx={{ fontWeight: 600 }}>
                        {formatNumber(debtTotals.outstandingBalance)}
                      </TableCell>
                      <TableCell align="right" sx={{ fontWeight: 600 }}>
                        {formatNumber(debtTotals.advancePayment)}
                      </TableCell>
                      <TableCell align="right" sx={{ fontWeight: 600 }}>
                        {formatNumber(debtTotals.accruedInterest)}
                      </TableCell>
                      <TableCell></TableCell>
                    </TableRow>
                  </TableBody>
                </Table>
              </TableContainer>
            </CardContent>
          </Card>
        </Grid>

        {/* Collateral Section */}
        <Grid item xs={12}>
          <Card>
            <CardContent>
              <SectionHeader icon={Home} title="담보물 정보" />
              <TableContainer component={Paper} variant="outlined">
                <Table size="small">
                  <TableHead>
                    <TableRow sx={{ bgcolor: 'grey.100' }}>
                      <TableCell sx={{ fontWeight: 600 }}>물건</TableCell>
                      <TableCell sx={{ fontWeight: 600 }}>물건유형</TableCell>
                      <TableCell sx={{ fontWeight: 600 }}>소재지</TableCell>
                      <TableCell sx={{ fontWeight: 600 }} align="right">토지면적(㎡)</TableCell>
                      <TableCell sx={{ fontWeight: 600 }} align="right">건물면적(㎡)</TableCell>
                      <TableCell sx={{ fontWeight: 600 }} align="right">감정평가액</TableCell>
                      <TableCell sx={{ fontWeight: 600 }} align="right">자체감정가</TableCell>
                      <TableCell sx={{ fontWeight: 600 }} align="right">낙찰가율</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {collateralInfo.map((col, index) => (
                      <TableRow key={col.id}>
                        <TableCell>물건 {index + 1}</TableCell>
                        <TableCell>{col.propertyType || '-'}</TableCell>
                        <TableCell sx={{ maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis' }}>
                          {col.address || '-'}
                        </TableCell>
                        <TableCell align="right">{formatNumber(col.landArea)}</TableCell>
                        <TableCell align="right">{formatNumber(col.buildingArea)}</TableCell>
                        <TableCell align="right">{formatNumber(col.appraisedValue)}</TableCell>
                        <TableCell align="right">{formatNumber(col.internalAppraisal)}</TableCell>
                        <TableCell align="right">{col.auctionRate}%</TableCell>
                      </TableRow>
                    ))}
                    <TableRow sx={{ bgcolor: 'primary.50' }}>
                      <TableCell colSpan={3} sx={{ fontWeight: 600 }}>합계</TableCell>
                      <TableCell align="right" sx={{ fontWeight: 600 }}>
                        {formatNumber(collateralTotals.landArea)}
                      </TableCell>
                      <TableCell align="right" sx={{ fontWeight: 600 }}>
                        {formatNumber(collateralTotals.buildingArea)}
                      </TableCell>
                      <TableCell align="right" sx={{ fontWeight: 600 }}>
                        {formatNumber(collateralTotals.appraisedValue)}
                      </TableCell>
                      <TableCell align="right" sx={{ fontWeight: 600 }}>
                        {formatNumber(collateralTotals.internalAppraisal)}
                      </TableCell>
                      <TableCell></TableCell>
                    </TableRow>
                  </TableBody>
                </Table>
              </TableContainer>
            </CardContent>
          </Card>
        </Grid>

        {/* Calculation Results Section */}
        <Grid item xs={12}>
          <Card sx={{ bgcolor: 'success.50', borderColor: 'success.200', borderWidth: 2, borderStyle: 'solid' }}>
            <CardContent>
              <SectionHeader icon={Assessment} title="평가 결과" color="success" />

              <Grid container spacing={3}>
                {/* Main Results */}
                <Grid item xs={12} md={6}>
                  <Paper variant="outlined" sx={{ p: 2 }}>
                    <Typography variant="subtitle2" sx={{ fontWeight: 600, mb: 2 }}>
                      주요 평가액
                    </Typography>
                    <Box sx={{ display: 'flex', flexDirection: 'column', gap: 1 }}>
                      <Box sx={{ display: 'flex', justifyContent: 'space-between' }}>
                        <Typography variant="body2" color="text.secondary">총 담보평가액</Typography>
                        <Typography variant="body1" sx={{ fontWeight: 600 }}>
                          {formatKRW(results.totalCollateralValue)}
                        </Typography>
                      </Box>
                      <Box sx={{ display: 'flex', justifyContent: 'space-between' }}>
                        <Typography variant="body2" color="text.secondary">총 회수시점채권액</Typography>
                        <Typography variant="body1" sx={{ fontWeight: 600 }}>
                          {formatKRW(results.totalRecoveryDebt)}
                        </Typography>
                      </Box>
                      <Divider sx={{ my: 1 }} />
                      <Box sx={{ display: 'flex', justifyContent: 'space-between' }}>
                        <Typography variant="body2" color="text.secondary">NRV (순회수금)</Typography>
                        <Typography variant="h6" sx={{ fontWeight: 700, color: 'primary.main' }}>
                          {formatKRW(results.totalNRV)}
                        </Typography>
                      </Box>
                      <Box sx={{ display: 'flex', justifyContent: 'space-between' }}>
                        <Typography variant="body2" color="text.secondary">NPV (매각대금)</Typography>
                        <Typography variant="h6" sx={{ fontWeight: 700, color: 'success.main' }}>
                          {formatKRW(results.totalNPV)}
                        </Typography>
                      </Box>
                      <Box sx={{ display: 'flex', justifyContent: 'space-between' }}>
                        <Typography variant="body2" color="text.secondary">FV (미래가)</Typography>
                        <Typography variant="h6" sx={{ fontWeight: 700 }}>
                          {formatKRW(results.totalFV)}
                        </Typography>
                      </Box>
                    </Box>
                  </Paper>
                </Grid>

                {/* Fee & Settlement */}
                <Grid item xs={12} md={6}>
                  <Paper variant="outlined" sx={{ p: 2 }}>
                    <Typography variant="subtitle2" sx={{ fontWeight: 600, mb: 2 }}>
                      비용 및 정산
                    </Typography>
                    <Box sx={{ display: 'flex', flexDirection: 'column', gap: 1 }}>
                      <Box sx={{ display: 'flex', justifyContent: 'space-between' }}>
                        <Typography variant="body2" color="text.secondary">총 수수료</Typography>
                        <Typography variant="body1">
                          {formatKRW(results.totalFees)}
                        </Typography>
                      </Box>
                      <Box sx={{ display: 'flex', justifyContent: 'space-between' }}>
                        <Typography variant="body2" color="text.secondary">정산차액</Typography>
                        <Typography
                          variant="body1"
                          sx={{
                            color: results.totalSettlement >= 0 ? 'success.main' : 'error.main',
                            fontWeight: 600,
                          }}
                        >
                          {formatKRW(results.totalSettlement)}
                        </Typography>
                      </Box>
                      <Divider sx={{ my: 1 }} />
                      <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <Typography variant="body2" color="text.secondary">OPB 대비 회수율</Typography>
                        <Chip
                          label={
                            results.totalOutstandingBalance > 0
                              ? formatPercent(results.totalNPV / results.totalOutstandingBalance)
                              : '-'
                          }
                          color="primary"
                          sx={{ fontWeight: 600 }}
                        />
                      </Box>
                      <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <Typography variant="body2" color="text.secondary">채권최고액 대비</Typography>
                        <Chip
                          label={
                            results.totalMaxDebtAmount > 0
                              ? formatPercent(results.totalNPV / results.totalMaxDebtAmount)
                              : '-'
                          }
                          color="secondary"
                          variant="outlined"
                        />
                      </Box>
                    </Box>
                  </Paper>
                </Grid>

                {/* Per-Property Breakdown */}
                {results.propertyResults && results.propertyResults.length > 1 && (
                  <Grid item xs={12}>
                    <Typography variant="subtitle2" sx={{ fontWeight: 600, mb: 1 }}>
                      물건별 상세
                    </Typography>
                    <TableContainer component={Paper} variant="outlined">
                      <Table size="small">
                        <TableHead>
                          <TableRow sx={{ bgcolor: 'grey.100' }}>
                            <TableCell sx={{ fontWeight: 600 }}>물건</TableCell>
                            <TableCell sx={{ fontWeight: 600 }} align="right">담보평가액</TableCell>
                            <TableCell sx={{ fontWeight: 600 }} align="right">NRV</TableCell>
                            <TableCell sx={{ fontWeight: 600 }} align="right">NPV</TableCell>
                            <TableCell sx={{ fontWeight: 600 }} align="right">FV</TableCell>
                            <TableCell sx={{ fontWeight: 600 }} align="right">수수료</TableCell>
                          </TableRow>
                        </TableHead>
                        <TableBody>
                          {results.propertyResults.map((prop, index) => (
                            <TableRow key={prop.propertyId || index}>
                              <TableCell>물건 {index + 1}</TableCell>
                              <TableCell align="right">{formatNumber(prop.collateralValue)}</TableCell>
                              <TableCell align="right">{formatNumber(prop.nrv)}</TableCell>
                              <TableCell align="right">{formatNumber(prop.npv)}</TableCell>
                              <TableCell align="right">{formatNumber(prop.fv)}</TableCell>
                              <TableCell align="right">{formatNumber(prop.totalFee)}</TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </TableContainer>
                  </Grid>
                )}
              </Grid>
            </CardContent>
          </Card>
        </Grid>
      </Grid>

      {/* Footer */}
      <Box sx={{ mt: 3, textAlign: 'center' }}>
        <Typography variant="caption" color="text.secondary">
          본 보고서는 {coverData.preparingOrg || 'MCI대부(주)'}에서 작성되었습니다.
        </Typography>
        <Typography variant="caption" display="block" color="text.secondary">
          * 금액 단위: 원 | 할인기간: {coverData.discountPeriod}일 | 현가할인율: {(coverData.discountRate * 100).toFixed(2)}%
        </Typography>
      </Box>
    </Box>
  )
}

export default EvaluationReportPreview
