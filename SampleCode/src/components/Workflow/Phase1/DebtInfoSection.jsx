import React from 'react'
import {
  Box,
  Card,
  CardContent,
  Typography,
  TextField,
  IconButton,
  Button,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  Tooltip,
  InputAdornment,
} from '@mui/material'
import { Add, Delete, AccountBalance } from '@mui/icons-material'
import useWorkflowStore from '../../../store/workflowStore'
import { formatNumber } from '../../../utils/nplCalculations'

const DebtInfoSection = () => {
  const { debtInfo, addDebtInfo, updateDebtInfo, removeDebtInfo } = useWorkflowStore()

  const handleChange = (id, field, value) => {
    // Convert number fields
    const numericFields = [
      'maxDebtAmount',
      'outstandingBalance',
      'advancePayment',
      'accruedInterest',
      'delinquencyRate',
    ]

    if (numericFields.includes(field)) {
      value = parseFloat(value) || 0
    }

    updateDebtInfo(id, { [field]: value })
  }

  const handleAddDebt = () => {
    addDebtInfo()
  }

  const handleRemoveDebt = (id) => {
    if (debtInfo.length > 1) {
      removeDebtInfo(id)
    }
  }

  // Calculate totals
  const totals = debtInfo.reduce(
    (acc, debt) => ({
      maxDebtAmount: acc.maxDebtAmount + (debt.maxDebtAmount || 0),
      outstandingBalance: acc.outstandingBalance + (debt.outstandingBalance || 0),
      advancePayment: acc.advancePayment + (debt.advancePayment || 0),
      accruedInterest: acc.accruedInterest + (debt.accruedInterest || 0),
    }),
    { maxDebtAmount: 0, outstandingBalance: 0, advancePayment: 0, accruedInterest: 0 }
  )

  return (
    <Card>
      <CardContent>
        <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 2 }}>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
            <AccountBalance color="primary" />
            <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
              채권현황
            </Typography>
            <Typography variant="body2" sx={{ color: 'text.secondary', ml: 1 }}>
              ({debtInfo.length}건)
            </Typography>
          </Box>
          <Button
            variant="outlined"
            size="small"
            startIcon={<Add />}
            onClick={handleAddDebt}
          >
            채권 추가
          </Button>
        </Box>

        <TableContainer component={Paper} variant="outlined">
          <Table size="small">
            <TableHead>
              <TableRow sx={{ bgcolor: 'grey.100' }}>
                <TableCell sx={{ fontWeight: 600, minWidth: 80 }}>순위</TableCell>
                <TableCell sx={{ fontWeight: 600, minWidth: 120 }}>금융기관</TableCell>
                <TableCell sx={{ fontWeight: 600, minWidth: 130 }}>계좌번호</TableCell>
                <TableCell sx={{ fontWeight: 600, minWidth: 130 }} align="right">채권최고액</TableCell>
                <TableCell sx={{ fontWeight: 600, minWidth: 130 }} align="right">대출잔액</TableCell>
                <TableCell sx={{ fontWeight: 600, minWidth: 110 }} align="right">가지급금</TableCell>
                <TableCell sx={{ fontWeight: 600, minWidth: 110 }} align="right">미수이자</TableCell>
                <TableCell sx={{ fontWeight: 600, minWidth: 100 }} align="right">연체이자율</TableCell>
                <TableCell sx={{ fontWeight: 600, minWidth: 110 }}>취급일</TableCell>
                <TableCell sx={{ fontWeight: 600, minWidth: 110 }}>만기일</TableCell>
                <TableCell sx={{ fontWeight: 600, minWidth: 50 }} align="center"></TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {debtInfo.map((debt, index) => (
                <TableRow key={debt.id} hover>
                  <TableCell>
                    <TextField
                      value={debt.priority}
                      onChange={(e) => handleChange(debt.id, 'priority', e.target.value)}
                      size="small"
                      variant="standard"
                      placeholder={`${index + 1}순위`}
                      sx={{ width: 70 }}
                    />
                  </TableCell>
                  <TableCell>
                    <TextField
                      value={debt.institutionName}
                      onChange={(e) => handleChange(debt.id, 'institutionName', e.target.value)}
                      size="small"
                      variant="standard"
                      placeholder="금융기관명"
                      sx={{ width: 110 }}
                    />
                  </TableCell>
                  <TableCell>
                    <TextField
                      value={debt.accountNumber}
                      onChange={(e) => handleChange(debt.id, 'accountNumber', e.target.value)}
                      size="small"
                      variant="standard"
                      placeholder="계좌번호"
                      sx={{ width: 120 }}
                    />
                  </TableCell>
                  <TableCell align="right">
                    <TextField
                      type="number"
                      value={debt.maxDebtAmount || ''}
                      onChange={(e) => handleChange(debt.id, 'maxDebtAmount', e.target.value)}
                      size="small"
                      variant="standard"
                      placeholder="0"
                      inputProps={{ style: { textAlign: 'right' } }}
                      sx={{ width: 120 }}
                    />
                  </TableCell>
                  <TableCell align="right">
                    <TextField
                      type="number"
                      value={debt.outstandingBalance || ''}
                      onChange={(e) => handleChange(debt.id, 'outstandingBalance', e.target.value)}
                      size="small"
                      variant="standard"
                      placeholder="0"
                      inputProps={{ style: { textAlign: 'right' } }}
                      sx={{ width: 120 }}
                    />
                  </TableCell>
                  <TableCell align="right">
                    <TextField
                      type="number"
                      value={debt.advancePayment || ''}
                      onChange={(e) => handleChange(debt.id, 'advancePayment', e.target.value)}
                      size="small"
                      variant="standard"
                      placeholder="0"
                      inputProps={{ style: { textAlign: 'right' } }}
                      sx={{ width: 100 }}
                    />
                  </TableCell>
                  <TableCell align="right">
                    <TextField
                      type="number"
                      value={debt.accruedInterest || ''}
                      onChange={(e) => handleChange(debt.id, 'accruedInterest', e.target.value)}
                      size="small"
                      variant="standard"
                      placeholder="0"
                      inputProps={{ style: { textAlign: 'right' } }}
                      sx={{ width: 100 }}
                    />
                  </TableCell>
                  <TableCell align="right">
                    <TextField
                      type="number"
                      value={debt.delinquencyRate || ''}
                      onChange={(e) => handleChange(debt.id, 'delinquencyRate', e.target.value)}
                      size="small"
                      variant="standard"
                      placeholder="0"
                      InputProps={{
                        endAdornment: <InputAdornment position="end">%</InputAdornment>,
                      }}
                      inputProps={{ style: { textAlign: 'right' }, step: 0.01 }}
                      sx={{ width: 80 }}
                    />
                  </TableCell>
                  <TableCell>
                    <TextField
                      type="date"
                      value={debt.handlingDate || ''}
                      onChange={(e) => handleChange(debt.id, 'handlingDate', e.target.value)}
                      size="small"
                      variant="standard"
                      InputLabelProps={{ shrink: true }}
                      sx={{ width: 100 }}
                    />
                  </TableCell>
                  <TableCell>
                    <TextField
                      type="date"
                      value={debt.maturityDate || ''}
                      onChange={(e) => handleChange(debt.id, 'maturityDate', e.target.value)}
                      size="small"
                      variant="standard"
                      InputLabelProps={{ shrink: true }}
                      sx={{ width: 100 }}
                    />
                  </TableCell>
                  <TableCell align="center">
                    <Tooltip title={debtInfo.length <= 1 ? '최소 1건 필요' : '삭제'}>
                      <span>
                        <IconButton
                          size="small"
                          onClick={() => handleRemoveDebt(debt.id)}
                          disabled={debtInfo.length <= 1}
                          color="error"
                        >
                          <Delete fontSize="small" />
                        </IconButton>
                      </span>
                    </Tooltip>
                  </TableCell>
                </TableRow>
              ))}

              {/* Totals Row */}
              <TableRow sx={{ bgcolor: 'primary.50' }}>
                <TableCell colSpan={3} sx={{ fontWeight: 600 }}>
                  합계
                </TableCell>
                <TableCell align="right" sx={{ fontWeight: 600 }}>
                  {formatNumber(totals.maxDebtAmount)}
                </TableCell>
                <TableCell align="right" sx={{ fontWeight: 600 }}>
                  {formatNumber(totals.outstandingBalance)}
                </TableCell>
                <TableCell align="right" sx={{ fontWeight: 600 }}>
                  {formatNumber(totals.advancePayment)}
                </TableCell>
                <TableCell align="right" sx={{ fontWeight: 600 }}>
                  {formatNumber(totals.accruedInterest)}
                </TableCell>
                <TableCell colSpan={4}></TableCell>
              </TableRow>
            </TableBody>
          </Table>
        </TableContainer>

        <Typography variant="caption" sx={{ display: 'block', mt: 1, color: 'text.secondary' }}>
          * 금액 단위: 원 | 이자율 단위: % (연율)
        </Typography>
      </CardContent>
    </Card>
  )
}

export default DebtInfoSection
