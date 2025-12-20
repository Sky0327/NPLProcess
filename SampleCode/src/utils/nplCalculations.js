/**
 * NPL (Non-Performing Loan) Calculation Engine
 *
 * Based on the evaluation report requirements:
 * - NPV (매각대금): Net Present Value calculation
 * - NRV (순회수금): Net Recovery Value
 * - FV (미래가): Future Value
 * - Settlement and fee calculations
 */

// Constants
const MORTGAGE_TRANSFER_RATE = 0.0048 // 0.48% - 근저당권이전비용률
const DEFAULT_MANAGEMENT_COST_RATE = 0.0095 // 0.95% - 관리비용률
const DEFAULT_DISCOUNT_RATE = 0.0628 // 6.28% - 현가할인율
const DEFAULT_DISCOUNT_PERIOD = 360 // days - 할인기간

/**
 * 담보평가액 = 감정평가액 × 낙찰가율
 * Collateral Value = Appraised Value × Auction Rate
 */
export const calculateCollateralValue = (appraisedValue, auctionRate) => {
  if (!appraisedValue || !auctionRate) return 0
  return Math.round(appraisedValue * (auctionRate / 100))
}

/**
 * 회수시점채권액 = 대출잔액 + 가지급금 + 미수이자 + (대출잔액 × 할인기간 × 연체이자율 / 365)
 * Recovery Debt = Outstanding Balance + Advance Payment + Accrued Interest + Daily Interest
 */
export const calculateRecoveryDebt = (debt, discountPeriod = DEFAULT_DISCOUNT_PERIOD) => {
  if (!debt) return 0

  const {
    outstandingBalance = 0,
    advancePayment = 0,
    accruedInterest = 0,
    delinquencyRate = 0,
  } = debt

  const dailyInterest = outstandingBalance * (discountPeriod / 365) * (delinquencyRate / 100)
  return Math.round(outstandingBalance + advancePayment + accruedInterest + dailyInterest)
}

/**
 * 총 회수시점채권액 - for multiple debts
 */
export const calculateTotalRecoveryDebt = (debtInfo, discountPeriod) => {
  if (!debtInfo || !Array.isArray(debtInfo)) return 0
  return debtInfo.reduce((sum, debt) => sum + calculateRecoveryDebt(debt, discountPeriod), 0)
}

/**
 * 근저당권이전비용 = 채권최고액 × 0.48%
 * Mortgage Transfer Cost = Max Debt Amount × 0.48%
 */
export const calculateMortgageTransferCost = (maxDebtAmount) => {
  if (!maxDebtAmount) return 0
  return Math.round(maxDebtAmount * MORTGAGE_TRANSFER_RATE)
}

/**
 * 총 근저당권이전비용 - for multiple debts
 */
export const calculateTotalMortgageTransferCost = (debtInfo) => {
  if (!debtInfo || !Array.isArray(debtInfo)) return 0
  return debtInfo.reduce((sum, debt) => sum + calculateMortgageTransferCost(debt.maxDebtAmount || 0), 0)
}

/**
 * 배당가능액 = 담보평가액 - 선순위채권액
 * Dividend Amount = Collateral Value - Senior Debt
 */
export const calculateDividendAmount = (collateralValue, seniorDebt = 0) => {
  return Math.max(0, collateralValue - seniorDebt)
}

/**
 * NRV(순회수금) = MAX(MIN(배당가능액, 근저당권설정액, 회수시점채권액) - 근저당권이전비용, 0)
 * Net Recovery Value = MAX(MIN(Dividend, Mortgage, Recovery Debt) - Transfer Cost, 0)
 */
export const calculateNRV = (dividendAmount, mortgageAmount, recoveryDebt, transferCost) => {
  if (!dividendAmount && !mortgageAmount && !recoveryDebt) return 0

  const minValue = Math.min(
    dividendAmount || Infinity,
    mortgageAmount || Infinity,
    recoveryDebt || Infinity
  )

  // If all values were 0 or undefined, minValue would be Infinity
  if (!isFinite(minValue)) return 0

  return Math.max(0, Math.round(minValue - (transferCost || 0)))
}

/**
 * 관리비용 = 순회수금 × 관리비용률
 * Management Cost = NRV × Management Cost Rate
 */
export const calculateManagementCost = (nrv, managementCostRate = DEFAULT_MANAGEMENT_COST_RATE) => {
  if (!nrv) return 0
  return Math.round(nrv * managementCostRate)
}

/**
 * NPV(매각대금) = (순회수금 - 관리비용) × (1 - 현가할인율 × n / 365)
 * NPV = (NRV - Management Cost) × (1 - Discount Rate × Days / 365)
 */
export const calculateNPV = (
  nrv,
  managementCost,
  discountRate = DEFAULT_DISCOUNT_RATE,
  discountPeriod = DEFAULT_DISCOUNT_PERIOD
) => {
  if (!nrv) return 0

  const netAmount = nrv - (managementCost || 0)
  const discountFactor = 1 - (discountRate * discountPeriod / 365)

  return Math.round(netAmount * discountFactor)
}

/**
 * FV(미래가) = 매각대금 × (1 + 현가할인율 × 할인기간 / 365)
 * Future Value = NPV × (1 + Discount Rate × Days / 365)
 */
export const calculateFV = (
  npv,
  discountRate = DEFAULT_DISCOUNT_RATE,
  discountPeriod = DEFAULT_DISCOUNT_PERIOD
) => {
  if (!npv) return 0

  const growthFactor = 1 + (discountRate * discountPeriod / 365)
  return Math.round(npv * growthFactor)
}

/**
 * 정산비용 = 근저당권이전비용 + 기타비용
 * Settlement Cost = Mortgage Transfer Cost + Other Costs
 */
export const calculateSettlementCost = (mortgageTransferCost, otherCosts = 0) => {
  return Math.round((mortgageTransferCost || 0) + otherCosts)
}

/**
 * 정산차액 = 확정배당금 - FV - 정산비용 - 관리비용
 * Settlement = Confirmed Dividend - FV - Settlement Cost - Management Cost
 */
export const calculateSettlement = (confirmedDividend, fv, settlementCost, managementCost) => {
  return Math.round(
    (confirmedDividend || 0) - (fv || 0) - (settlementCost || 0) - (managementCost || 0)
  )
}

/**
 * 총수수료 = FV - NPV + 관리비용
 * Total Fee = FV - NPV + Management Cost
 */
export const calculateTotalFee = (fv, npv, managementCost) => {
  return Math.round((fv || 0) - (npv || 0) + (managementCost || 0))
}

/**
 * 현가할인수수료 = 매각대금 × 현가할인율 × (할인기간 / 365)
 * Present Value Discount Fee = Sale Price × Discount Rate × (Days / 365)
 */
export const calculatePVDiscountFee = (
  salePrice,
  discountRate = DEFAULT_DISCOUNT_RATE,
  discountPeriod = DEFAULT_DISCOUNT_PERIOD
) => {
  if (!salePrice) return 0
  return Math.round(salePrice * discountRate * (discountPeriod / 365))
}

/**
 * OPB 대비 회수율 = NPV / (대출잔액 + 가지급금)
 * Recovery Rate vs OPB = NPV / (Outstanding Balance + Advance Payment)
 */
export const calculateRecoveryRate = (npv, outstandingBalance, advancePayment = 0) => {
  const totalDebt = (outstandingBalance || 0) + (advancePayment || 0)
  if (!totalDebt) return 0
  return (npv || 0) / totalDebt
}

/**
 * Calculate all results for a single property/debt combination
 */
export const calculatePropertyResults = (
  collateral,
  debtInfo,
  coverData = {}
) => {
  const {
    discountRate = DEFAULT_DISCOUNT_RATE,
    discountPeriod = DEFAULT_DISCOUNT_PERIOD,
    managementCostRate = DEFAULT_MANAGEMENT_COST_RATE,
  } = coverData

  // Get values from collateral
  const appraisedValue = collateral?.appraisedValue || collateral?.internalAppraisal || 0
  const auctionRate = collateral?.auctionRate || 80 // Default 80%

  // Calculate collateral value
  const collateralValue = calculateCollateralValue(appraisedValue, auctionRate)

  // Calculate debt-related values
  const totalMaxDebtAmount = debtInfo.reduce((sum, d) => sum + (d.maxDebtAmount || 0), 0)
  const totalOutstandingBalance = debtInfo.reduce((sum, d) => sum + (d.outstandingBalance || 0), 0)
  const totalAdvancePayment = debtInfo.reduce((sum, d) => sum + (d.advancePayment || 0), 0)

  const recoveryDebt = calculateTotalRecoveryDebt(debtInfo, discountPeriod)
  const mortgageTransferCost = calculateTotalMortgageTransferCost(debtInfo)

  // Calculate NRV (assuming no senior debt for simplicity - can be extended)
  const dividendAmount = calculateDividendAmount(collateralValue, 0)
  const nrv = calculateNRV(dividendAmount, totalMaxDebtAmount, recoveryDebt, mortgageTransferCost)

  // Calculate costs and values
  const managementCost = calculateManagementCost(nrv, managementCostRate)
  const npv = calculateNPV(nrv, managementCost, discountRate, discountPeriod)
  const fv = calculateFV(npv, discountRate, discountPeriod)

  const settlementCost = calculateSettlementCost(mortgageTransferCost)
  const settlement = calculateSettlement(nrv, fv, settlementCost, managementCost)
  const totalFee = calculateTotalFee(fv, npv, managementCost)

  const pvDiscountFee = calculatePVDiscountFee(npv, discountRate, discountPeriod)
  const recoveryRate = calculateRecoveryRate(npv, totalOutstandingBalance, totalAdvancePayment)

  return {
    // Input values
    appraisedValue,
    auctionRate,

    // Calculated values
    collateralValue,
    recoveryDebt,
    mortgageTransferCost,
    dividendAmount,
    nrv,
    managementCost,
    npv,
    fv,
    settlementCost,
    settlement,
    totalFee,
    pvDiscountFee,
    recoveryRate,

    // Debt totals
    totalMaxDebtAmount,
    totalOutstandingBalance,
    totalAdvancePayment,
  }
}

/**
 * Calculate all results for multiple properties
 * Returns aggregated totals and per-property breakdowns
 */
export const calculateAllResults = (collateralInfo, debtInfo, coverData) => {
  if (!collateralInfo || !Array.isArray(collateralInfo) || collateralInfo.length === 0) {
    return {
      propertyResults: [],
      totalCollateralValue: 0,
      totalNRV: 0,
      totalNPV: 0,
      totalFV: 0,
      totalSettlement: 0,
      totalFees: 0,
      totalMaxDebtAmount: 0,
      totalOutstandingBalance: 0,
      totalRecoveryDebt: 0,
    }
  }

  // Calculate results for each property
  const propertyResults = collateralInfo.map((collateral) => ({
    propertyId: collateral.id,
    propertyIndex: collateral.propertyIndex,
    address: collateral.address,
    ...calculatePropertyResults(collateral, debtInfo, coverData),
  }))

  // Aggregate totals
  const totals = propertyResults.reduce(
    (acc, result) => ({
      totalCollateralValue: acc.totalCollateralValue + result.collateralValue,
      totalNRV: acc.totalNRV + result.nrv,
      totalNPV: acc.totalNPV + result.npv,
      totalFV: acc.totalFV + result.fv,
      totalSettlement: acc.totalSettlement + result.settlement,
      totalFees: acc.totalFees + result.totalFee,
    }),
    {
      totalCollateralValue: 0,
      totalNRV: 0,
      totalNPV: 0,
      totalFV: 0,
      totalSettlement: 0,
      totalFees: 0,
    }
  )

  // Debt totals (calculated once, not per property)
  const {
    discountPeriod = DEFAULT_DISCOUNT_PERIOD,
  } = coverData || {}

  const totalMaxDebtAmount = debtInfo?.reduce((sum, d) => sum + (d.maxDebtAmount || 0), 0) || 0
  const totalOutstandingBalance = debtInfo?.reduce((sum, d) => sum + (d.outstandingBalance || 0), 0) || 0
  const totalRecoveryDebt = calculateTotalRecoveryDebt(debtInfo, discountPeriod)

  return {
    propertyResults,
    ...totals,
    totalMaxDebtAmount,
    totalOutstandingBalance,
    totalRecoveryDebt,
  }
}

/**
 * Format currency in Korean Won
 */
export const formatKRW = (value) => {
  if (value === null || value === undefined) return '-'
  return new Intl.NumberFormat('ko-KR', {
    style: 'currency',
    currency: 'KRW',
    maximumFractionDigits: 0,
  }).format(value)
}

/**
 * Format number with commas
 */
export const formatNumber = (value) => {
  if (value === null || value === undefined) return '-'
  return new Intl.NumberFormat('ko-KR').format(value)
}

/**
 * Format percentage
 */
export const formatPercent = (value, decimals = 2) => {
  if (value === null || value === undefined) return '-'
  return `${(value * 100).toFixed(decimals)}%`
}

export default {
  calculateCollateralValue,
  calculateRecoveryDebt,
  calculateTotalRecoveryDebt,
  calculateMortgageTransferCost,
  calculateTotalMortgageTransferCost,
  calculateDividendAmount,
  calculateNRV,
  calculateManagementCost,
  calculateNPV,
  calculateFV,
  calculateSettlementCost,
  calculateSettlement,
  calculateTotalFee,
  calculatePVDiscountFee,
  calculateRecoveryRate,
  calculatePropertyResults,
  calculateAllResults,
  formatKRW,
  formatNumber,
  formatPercent,
}
