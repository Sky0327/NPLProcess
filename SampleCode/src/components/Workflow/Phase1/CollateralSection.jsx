import React from 'react'
import {
  Box,
  Card,
  CardContent,
  Typography,
  TextField,
  Grid,
  IconButton,
  Button,
  MenuItem,
  Divider,
  Collapse,
  InputAdornment,
} from '@mui/material'
import {
  Add,
  Delete,
  Home,
  ExpandMore,
  ExpandLess,
  Landscape,
  Apartment,
} from '@mui/icons-material'
import useWorkflowStore from '../../../store/workflowStore'
import { formatNumber } from '../../../utils/nplCalculations'

const PROPERTY_TYPES = [
  '아파트',
  '오피스텔',
  '다세대주택',
  '다가구주택',
  '단독주택',
  '연립주택',
  '상가',
  '근린생활시설',
  '토지',
  '공장',
  '창고',
  '기타',
]

const LAND_USE_ZONES = [
  '주거지역',
  '상업지역',
  '공업지역',
  '녹지지역',
  '관리지역',
  '농림지역',
  '자연환경보전지역',
  '기타',
]

const CollateralItem = ({ collateral, onUpdate, onRemove, canRemove, index }) => {
  const [expanded, setExpanded] = React.useState(true)

  const handleChange = (field, value) => {
    // Convert number fields
    const numericFields = [
      'landArea',
      'landPrice',
      'buildingArea',
      'appraisedValue',
      'internalAppraisal',
      'auctionRate',
    ]

    if (numericFields.includes(field)) {
      value = parseFloat(value) || 0
    }

    onUpdate(collateral.id, { [field]: value })
  }

  return (
    <Card variant="outlined" sx={{ mb: 2 }}>
      <CardContent sx={{ pb: expanded ? 2 : 1 }}>
        {/* Header */}
        <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
            <Home color="primary" fontSize="small" />
            <Typography variant="subtitle2" sx={{ fontWeight: 600 }}>
              물건 {index + 1}
            </Typography>
            {collateral.propertyType && (
              <Typography variant="body2" sx={{ color: 'text.secondary' }}>
                - {collateral.propertyType}
              </Typography>
            )}
            {collateral.address && (
              <Typography variant="body2" sx={{ color: 'text.secondary', ml: 1 }}>
                ({collateral.address.substring(0, 30)}...)
              </Typography>
            )}
          </Box>
          <Box>
            <IconButton size="small" onClick={() => setExpanded(!expanded)}>
              {expanded ? <ExpandLess /> : <ExpandMore />}
            </IconButton>
            <IconButton
              size="small"
              onClick={() => onRemove(collateral.id)}
              disabled={!canRemove}
              color="error"
            >
              <Delete fontSize="small" />
            </IconButton>
          </Box>
        </Box>

        <Collapse in={expanded}>
          <Divider sx={{ my: 2 }} />

          {/* 기본 정보 */}
          <Grid container spacing={2}>
            <Grid item xs={12} sm={6} md={4}>
              <TextField
                select
                label="물건유형"
                value={collateral.propertyType}
                onChange={(e) => handleChange('propertyType', e.target.value)}
                fullWidth
                size="small"
              >
                {PROPERTY_TYPES.map((option) => (
                  <MenuItem key={option} value={option}>
                    {option}
                  </MenuItem>
                ))}
              </TextField>
            </Grid>
            <Grid item xs={12} sm={6} md={8}>
              <TextField
                label="소재지"
                value={collateral.address}
                onChange={(e) => handleChange('address', e.target.value)}
                fullWidth
                size="small"
                placeholder="예: 서울 관악구 신림동 1694 신림현대아파트 제103동 제303호"
              />
            </Grid>
          </Grid>

          {/* 토지 정보 */}
          <Box sx={{ mt: 3 }}>
            <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 1 }}>
              <Landscape fontSize="small" color="action" />
              <Typography variant="body2" sx={{ fontWeight: 600 }}>
                토지 정보
              </Typography>
            </Box>
            <Grid container spacing={2}>
              <Grid item xs={12} sm={4}>
                <TextField
                  select
                  label="용도지역"
                  value={collateral.landUseZone}
                  onChange={(e) => handleChange('landUseZone', e.target.value)}
                  fullWidth
                  size="small"
                >
                  {LAND_USE_ZONES.map((option) => (
                    <MenuItem key={option} value={option}>
                      {option}
                    </MenuItem>
                  ))}
                </TextField>
              </Grid>
              <Grid item xs={12} sm={4}>
                <TextField
                  label="면적"
                  type="number"
                  value={collateral.landArea || ''}
                  onChange={(e) => handleChange('landArea', e.target.value)}
                  fullWidth
                  size="small"
                  InputProps={{
                    endAdornment: <InputAdornment position="end">㎡</InputAdornment>,
                  }}
                />
              </Grid>
              <Grid item xs={12} sm={4}>
                <TextField
                  label="개별공시지가"
                  type="number"
                  value={collateral.landPrice || ''}
                  onChange={(e) => handleChange('landPrice', e.target.value)}
                  fullWidth
                  size="small"
                  InputProps={{
                    endAdornment: <InputAdornment position="end">원/㎡</InputAdornment>,
                  }}
                />
              </Grid>
            </Grid>
          </Box>

          {/* 건물 정보 */}
          <Box sx={{ mt: 3 }}>
            <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 1 }}>
              <Apartment fontSize="small" color="action" />
              <Typography variant="body2" sx={{ fontWeight: 600 }}>
                건물 정보
              </Typography>
            </Box>
            <Grid container spacing={2}>
              <Grid item xs={12} sm={4}>
                <TextField
                  label="면적"
                  type="number"
                  value={collateral.buildingArea || ''}
                  onChange={(e) => handleChange('buildingArea', e.target.value)}
                  fullWidth
                  size="small"
                  InputProps={{
                    endAdornment: <InputAdornment position="end">㎡</InputAdornment>,
                  }}
                />
              </Grid>
              <Grid item xs={12} sm={4}>
                <TextField
                  label="규모"
                  value={collateral.buildingScale}
                  onChange={(e) => handleChange('buildingScale', e.target.value)}
                  fullWidth
                  size="small"
                  placeholder="예: 지하1층/지상15층"
                />
              </Grid>
              <Grid item xs={12} sm={4}>
                <TextField
                  label="구조"
                  value={collateral.buildingStructure}
                  onChange={(e) => handleChange('buildingStructure', e.target.value)}
                  fullWidth
                  size="small"
                  placeholder="예: 철근콘크리트조"
                />
              </Grid>
            </Grid>
          </Box>

          {/* 감정평가 정보 */}
          <Box sx={{ mt: 3 }}>
            <Typography variant="body2" sx={{ fontWeight: 600, mb: 1 }}>
              감정평가 정보
            </Typography>
            <Grid container spacing={2}>
              <Grid item xs={12} sm={6} md={3}>
                <TextField
                  label="대출당시 감정평가액"
                  type="number"
                  value={collateral.appraisedValue || ''}
                  onChange={(e) => handleChange('appraisedValue', e.target.value)}
                  fullWidth
                  size="small"
                  InputProps={{
                    endAdornment: <InputAdornment position="end">원</InputAdornment>,
                  }}
                />
              </Grid>
              <Grid item xs={12} sm={6} md={3}>
                <TextField
                  label="자체감정가"
                  type="number"
                  value={collateral.internalAppraisal || ''}
                  onChange={(e) => handleChange('internalAppraisal', e.target.value)}
                  fullWidth
                  size="small"
                  InputProps={{
                    endAdornment: <InputAdornment position="end">원</InputAdornment>,
                  }}
                />
              </Grid>
              <Grid item xs={12} sm={6} md={3}>
                <TextField
                  label="평가기관"
                  value={collateral.appraiser}
                  onChange={(e) => handleChange('appraiser', e.target.value)}
                  fullWidth
                  size="small"
                  placeholder="예: 한국감정원"
                />
              </Grid>
              <Grid item xs={12} sm={6} md={3}>
                <TextField
                  label="낙찰가율"
                  type="number"
                  value={collateral.auctionRate || ''}
                  onChange={(e) => handleChange('auctionRate', e.target.value)}
                  fullWidth
                  size="small"
                  InputProps={{
                    endAdornment: <InputAdornment position="end">%</InputAdornment>,
                  }}
                  inputProps={{ step: 0.1 }}
                  helperText="예상 낙찰가율"
                />
              </Grid>
            </Grid>
          </Box>
        </Collapse>
      </CardContent>
    </Card>
  )
}

const CollateralSection = () => {
  const {
    collateralInfo,
    addCollateralInfo,
    updateCollateralInfo,
    removeCollateralInfo,
  } = useWorkflowStore()

  const handleAddCollateral = () => {
    addCollateralInfo()
  }

  // Calculate totals
  const totals = collateralInfo.reduce(
    (acc, col) => ({
      appraisedValue: acc.appraisedValue + (col.appraisedValue || 0),
      internalAppraisal: acc.internalAppraisal + (col.internalAppraisal || 0),
      buildingArea: acc.buildingArea + (col.buildingArea || 0),
      landArea: acc.landArea + (col.landArea || 0),
    }),
    { appraisedValue: 0, internalAppraisal: 0, buildingArea: 0, landArea: 0 }
  )

  return (
    <Box>
      <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 2 }}>
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
          <Home color="primary" />
          <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
            담보물 정보
          </Typography>
          <Typography variant="body2" sx={{ color: 'text.secondary', ml: 1 }}>
            ({collateralInfo.length}건)
          </Typography>
        </Box>
        <Button
          variant="outlined"
          size="small"
          startIcon={<Add />}
          onClick={handleAddCollateral}
        >
          물건 추가
        </Button>
      </Box>

      {/* Collateral Items */}
      {collateralInfo.map((collateral, index) => (
        <CollateralItem
          key={collateral.id}
          collateral={collateral}
          onUpdate={updateCollateralInfo}
          onRemove={removeCollateralInfo}
          canRemove={collateralInfo.length > 1}
          index={index}
        />
      ))}

      {/* Summary Card */}
      <Card sx={{ bgcolor: 'grey.50' }}>
        <CardContent>
          <Typography variant="subtitle2" sx={{ fontWeight: 600, mb: 2 }}>
            합계
          </Typography>
          <Grid container spacing={2}>
            <Grid item xs={6} sm={3}>
              <Typography variant="caption" color="text.secondary">
                총 감정평가액
              </Typography>
              <Typography variant="body1" sx={{ fontWeight: 600 }}>
                {formatNumber(totals.appraisedValue)}원
              </Typography>
            </Grid>
            <Grid item xs={6} sm={3}>
              <Typography variant="caption" color="text.secondary">
                총 자체감정가
              </Typography>
              <Typography variant="body1" sx={{ fontWeight: 600 }}>
                {formatNumber(totals.internalAppraisal)}원
              </Typography>
            </Grid>
            <Grid item xs={6} sm={3}>
              <Typography variant="caption" color="text.secondary">
                총 건물면적
              </Typography>
              <Typography variant="body1" sx={{ fontWeight: 600 }}>
                {formatNumber(totals.buildingArea)}㎡
              </Typography>
            </Grid>
            <Grid item xs={6} sm={3}>
              <Typography variant="caption" color="text.secondary">
                총 토지면적
              </Typography>
              <Typography variant="body1" sx={{ fontWeight: 600 }}>
                {formatNumber(totals.landArea)}㎡
              </Typography>
            </Grid>
          </Grid>
        </CardContent>
      </Card>
    </Box>
  )
}

export default CollateralSection
