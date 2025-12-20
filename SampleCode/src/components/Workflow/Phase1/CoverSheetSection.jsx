import React from 'react'
import {
  Box,
  Card,
  CardContent,
  Typography,
  TextField,
  Grid,
  MenuItem,
  Divider,
  InputAdornment,
} from '@mui/material'
import { Business, Person, Home, Calculate } from '@mui/icons-material'
import useWorkflowStore from '../../../store/workflowStore'

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

const BUSINESS_CLASSIFICATIONS = ['개인', '기업']

const DISCOUNT_PERIODS = [
  { value: 360, label: '360일 (경매개시)' },
  { value: 420, label: '420일 (경매미개시/신탁)' },
]

const CoverSheetSection = () => {
  const { coverData, setCoverData, projectConfig, setProjectConfig } = useWorkflowStore()

  const handleCoverChange = (field) => (event) => {
    setCoverData({ [field]: event.target.value })
  }

  const handleProjectChange = (field) => (event) => {
    setProjectConfig({ [field]: event.target.value })
  }

  const handleRateChange = (field) => (event) => {
    const value = parseFloat(event.target.value) / 100
    setCoverData({ [field]: isNaN(value) ? 0 : value })
  }

  const handlePeriodChange = (event) => {
    setCoverData({ discountPeriod: event.target.value })
  }

  return (
    <Box sx={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
      {/* 프로젝트 기본 정보 */}
      <Card>
        <CardContent>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 2 }}>
            <Business color="primary" />
            <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
              프로젝트 정보
            </Typography>
          </Box>

          <Grid container spacing={2}>
            <Grid item xs={12} sm={6}>
              <TextField
                label="보고서명"
                value={projectConfig.reportName}
                onChange={handleProjectChange('reportName')}
                fullWidth
                size="small"
                required
                placeholder="예: 2024년 1분기 NPL 평가"
              />
            </Grid>
            <Grid item xs={12} sm={6}>
              <TextField
                label="프로젝트 ID"
                value={projectConfig.projectId}
                onChange={handleProjectChange('projectId')}
                fullWidth
                size="small"
                required
                placeholder="예: NPL-2024-001"
              />
            </Grid>
          </Grid>
        </CardContent>
      </Card>

      {/* Cover Sheet 기본 정보 */}
      <Card>
        <CardContent>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 2 }}>
            <Person color="primary" />
            <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
              Cover Sheet - 기본 정보
            </Typography>
          </Box>

          <Grid container spacing={2}>
            <Grid item xs={12} sm={6}>
              <TextField
                label="매각기관"
                value={coverData.sellingInstitution}
                onChange={handleCoverChange('sellingInstitution')}
                fullWidth
                size="small"
                required
                placeholder="예: 임오새마을금고"
              />
            </Grid>
            <Grid item xs={12} sm={6}>
              <TextField
                label="업무구분"
                value={coverData.businessType}
                onChange={handleCoverChange('businessType')}
                fullWidth
                size="small"
                placeholder="사후재정산방식 담보부채권 양수도"
              />
            </Grid>
            <Grid item xs={12} sm={6}>
              <TextField
                label="차주명"
                value={coverData.borrowerName}
                onChange={handleCoverChange('borrowerName')}
                fullWidth
                size="small"
                required
                placeholder="예: 김OO"
              />
            </Grid>
            <Grid item xs={12} sm={6}>
              <TextField
                select
                label="사업자구분"
                value={coverData.businessClassification}
                onChange={handleCoverChange('businessClassification')}
                fullWidth
                size="small"
              >
                {BUSINESS_CLASSIFICATIONS.map((option) => (
                  <MenuItem key={option} value={option}>
                    {option}
                  </MenuItem>
                ))}
              </TextField>
            </Grid>
            <Grid item xs={12} sm={6}>
              <TextField
                select
                label="물건유형"
                value={coverData.propertyType}
                onChange={handleCoverChange('propertyType')}
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
            <Grid item xs={12} sm={6}>
              <TextField
                label="작성기관"
                value={coverData.preparingOrg}
                onChange={handleCoverChange('preparingOrg')}
                fullWidth
                size="small"
              />
            </Grid>
            <Grid item xs={12}>
              <TextField
                label="소재지"
                value={coverData.address}
                onChange={handleCoverChange('address')}
                fullWidth
                size="small"
                required
                placeholder="예: 서울 관악구 신림동 1694 신림현대아파트 제103동 제303호"
                helperText="2단계 데이터 조회에 필요합니다"
              />
            </Grid>
            <Grid item xs={12} sm={6}>
              <TextField
                label="작성일"
                type="date"
                value={coverData.reportDate || ''}
                onChange={handleCoverChange('reportDate')}
                fullWidth
                size="small"
                InputLabelProps={{ shrink: true }}
              />
            </Grid>
          </Grid>
        </CardContent>
      </Card>

      {/* 주요 가정 */}
      <Card>
        <CardContent>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 2 }}>
            <Calculate color="primary" />
            <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
              주요 가정
            </Typography>
          </Box>

          <Grid container spacing={2}>
            <Grid item xs={12} sm={6} md={4}>
              <TextField
                label="기본할인율"
                type="number"
                value={(coverData.baseRate * 100).toFixed(2)}
                onChange={handleRateChange('baseRate')}
                fullWidth
                size="small"
                InputProps={{
                  endAdornment: <InputAdornment position="end">%</InputAdornment>,
                }}
                inputProps={{ step: 0.01 }}
                helperText="차입이자율"
              />
            </Grid>
            <Grid item xs={12} sm={6} md={4}>
              <TextField
                label="매입할인율"
                type="number"
                value={(coverData.purchaseRate * 100).toFixed(2)}
                onChange={handleRateChange('purchaseRate')}
                fullWidth
                size="small"
                InputProps={{
                  endAdornment: <InputAdornment position="end">%</InputAdornment>,
                }}
                inputProps={{ step: 0.01 }}
                helperText="고정 0.88%"
              />
            </Grid>
            <Grid item xs={12} sm={6} md={4}>
              <TextField
                label="현가할인율"
                type="number"
                value={(coverData.discountRate * 100).toFixed(2)}
                onChange={handleRateChange('discountRate')}
                fullWidth
                size="small"
                InputProps={{
                  endAdornment: <InputAdornment position="end">%</InputAdornment>,
                }}
                inputProps={{ step: 0.01 }}
                helperText="기본할인율 + 매입할인율"
              />
            </Grid>
            <Grid item xs={12} sm={6} md={4}>
              <TextField
                label="관리비용률"
                type="number"
                value={(coverData.managementCostRate * 100).toFixed(2)}
                onChange={handleRateChange('managementCostRate')}
                fullWidth
                size="small"
                InputProps={{
                  endAdornment: <InputAdornment position="end">%</InputAdornment>,
                }}
                inputProps={{ step: 0.01 }}
                helperText="총 회수금액의 0.95%"
              />
            </Grid>
            <Grid item xs={12} sm={6} md={4}>
              <TextField
                select
                label="할인기간"
                value={coverData.discountPeriod}
                onChange={handlePeriodChange}
                fullWidth
                size="small"
                helperText="경매개시 360일 / 경매미개시 420일"
              >
                {DISCOUNT_PERIODS.map((option) => (
                  <MenuItem key={option.value} value={option.value}>
                    {option.label}
                  </MenuItem>
                ))}
              </TextField>
            </Grid>
          </Grid>
        </CardContent>
      </Card>

      {/* API 설정 */}
      <Card>
        <CardContent>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 2 }}>
            <Home color="primary" />
            <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
              API 인증 설정
            </Typography>
          </Box>

          <Grid container spacing={2}>
            <Grid item xs={12} sm={6}>
              <TextField
                label="API ID"
                value={projectConfig.apiId}
                onChange={handleProjectChange('apiId')}
                fullWidth
                size="small"
                placeholder="API 사용자 ID"
              />
            </Grid>
            <Grid item xs={12} sm={6}>
              <TextField
                label="API Password"
                type="password"
                value={projectConfig.apiPassword}
                onChange={handleProjectChange('apiPassword')}
                fullWidth
                size="small"
                placeholder="API 비밀번호"
              />
            </Grid>
            <Grid item xs={12}>
              <TextField
                label="입력 폴더 경로"
                value={projectConfig.inputFolderPath}
                onChange={handleProjectChange('inputFolderPath')}
                fullWidth
                size="small"
                placeholder="예: C:\NPLData\Input"
                helperText="등기목록 및 Input 데이터 폴더 경로"
              />
            </Grid>
          </Grid>
        </CardContent>
      </Card>
    </Box>
  )
}

export default CoverSheetSection
