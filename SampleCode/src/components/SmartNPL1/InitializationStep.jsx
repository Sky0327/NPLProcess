import React, { useState } from 'react'
import {
  Box,
  TextField,
  Button,
  Card,
  CardContent,
  Typography,
  Grid,
  Divider,
} from '@mui/material'
import { Save, Settings } from '@mui/icons-material'

const InitializationStep = ({ projectData, setProjectData, onNext }) => {
  const [formData, setFormData] = useState({
    reportName: projectData.reportName || '',
    projectId: projectData.projectId || '',
    apiId: '',
    apiPassword: '',
    inputFolderPath: '',
  })

  const handleChange = (field) => (event) => {
    setFormData({
      ...formData,
      [field]: event.target.value,
    })
  }

  const handleSave = () => {
    setProjectData({
      ...projectData,
      reportName: formData.reportName,
      projectId: formData.projectId,
      apiSettings: {
        id: formData.apiId,
        password: formData.apiPassword,
      },
      inputFolderPath: formData.inputFolderPath,
    })
    alert('설정이 저장되었습니다.')
  }

  const handleInitialize = () => {
    if (!formData.reportName || !formData.projectId) {
      alert('보고서명과 프로젝트 ID를 입력해주세요.')
      return
    }
    handleSave()
    onNext()
  }

  return (
    <Card>
      <CardContent>
        <Box sx={{ display: 'flex', alignItems: 'center', mb: 3 }}>
          <Settings sx={{ mr: 1, color: 'primary.main' }} />
          <Typography variant="h5" component="h2">
            초기화 및 설정
          </Typography>
        </Box>

        <Divider sx={{ mb: 3 }} />

        <Grid container spacing={3}>
          <Grid item xs={12} md={6}>
            <TextField
              fullWidth
              label="보고서명"
              value={formData.reportName}
              onChange={handleChange('reportName')}
              required
              placeholder="예: 2024년 1분기 NPL 평가"
            />
          </Grid>
          <Grid item xs={12} md={6}>
            <TextField
              fullWidth
              label="프로젝트 ID"
              value={formData.projectId}
              onChange={handleChange('projectId')}
              required
              placeholder="예: PROJ-2024-001"
            />
          </Grid>

          <Grid item xs={12}>
            <Divider sx={{ my: 2 }} />
            <Typography variant="h6" gutterBottom>
              API 설정
            </Typography>
          </Grid>

          <Grid item xs={12} md={6}>
            <TextField
              fullWidth
              label="API ID"
              type="password"
              value={formData.apiId}
              onChange={handleChange('apiId')}
              placeholder="API 인증 ID"
            />
          </Grid>
          <Grid item xs={12} md={6}>
            <TextField
              fullWidth
              label="API Password"
              type="password"
              value={formData.apiPassword}
              onChange={handleChange('apiPassword')}
              placeholder="API 인증 비밀번호"
            />
          </Grid>

          <Grid item xs={12}>
            <TextField
              fullWidth
              label="입력 폴더 경로"
              value={formData.inputFolderPath}
              onChange={handleChange('inputFolderPath')}
              placeholder="등기목록 및 Input 데이터 폴더 경로"
              helperText="등기목록과 Input 데이터가 저장된 폴더 경로를 입력하세요"
            />
          </Grid>

          <Grid item xs={12}>
            <Box sx={{ display: 'flex', gap: 2, justifyContent: 'flex-end', mt: 3 }}>
              <Button
                variant="outlined"
                startIcon={<Save />}
                onClick={handleSave}
              >
                설정 저장
              </Button>
              <Button
                variant="contained"
                onClick={handleInitialize}
                disabled={!formData.reportName || !formData.projectId}
              >
                초기화 완료 및 다음 단계
              </Button>
            </Box>
          </Grid>
        </Grid>
      </CardContent>
    </Card>
  )
}

export default InitializationStep

