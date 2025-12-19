import React, { useState } from 'react'
import {
  Box,
  Button,
  Card,
  CardContent,
  Typography,
  Grid,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  Chip,
} from '@mui/material'
import {
  Description,
  Download,
  CheckCircle,
} from '@mui/icons-material'

const reports = [
  {
    id: 'report-2-1',
    name: '[2-1] 담보물정보',
    dependencies: ['registry', 'kbPrice'],
    description: '담보물건 현황 리포트',
  },
  {
    id: 'report-2-2',
    name: '[2-2] 감정평가',
    dependencies: ['kbPrice'],
    description: 'KB 부동산 정보 기반 감정평가',
  },
  {
    id: 'report-3',
    name: '[3] 경매정보',
    dependencies: ['auction'],
    description: '법원 경매 정보 리포트',
  },
  {
    id: 'report-5-1',
    name: '[5-1] 낙찰통계',
    dependencies: ['infocareStats'],
    description: '인포케어 통계 기반 낙찰 통계',
  },
  {
    id: 'report-5-2',
    name: '[5-2] 낙찰사례',
    dependencies: ['infocareCases'],
    description: '인포케어 사례 기반 낙찰 사례',
  },
  {
    id: 'report-6-1-gookto',
    name: '[6-1] 실거래사례_국토',
    dependencies: ['distanceGookto'],
    description: '국토부 실거래가 기반 사례',
  },
  {
    id: 'report-6-1-valuemap',
    name: '[6-1] 실거래사례_밸류맵',
    dependencies: ['distanceValuemap'],
    description: '밸류맵 실거래가 기반 사례',
  },
]

const ReportGenerationStep = ({
  projectData,
  inquiryResults,
  processingResults,
  onBack,
}) => {
  const [generatedReports, setGeneratedReports] = useState([])

  const checkDependencies = (report) => {
    return report.dependencies.every((dep) => {
      return inquiryResults[dep] || processingResults[dep]
    })
  }

  const handleGenerateReport = (report) => {
    if (!checkDependencies(report)) {
      alert(`필수 데이터가 준비되지 않았습니다: ${report.dependencies.join(', ')}`)
      return
    }

    setGeneratedReports([...generatedReports, report.id])
    alert(`${report.name} 리포트가 생성되었습니다.`)
  }

  const handleExport = () => {
    alert('리포트가 Excel 파일로 내보내기 되었습니다.')
  }

  return (
    <Card>
      <CardContent>
        <Box sx={{ display: 'flex', justifyContent: 'space-between', mb: 3 }}>
          <Box>
            <Typography variant="h5" component="h2" gutterBottom>
              리포트 생성
            </Typography>
            <Typography variant="body2" color="text.secondary">
              수집 및 가공된 데이터를 기반으로 리포트를 생성합니다.
            </Typography>
          </Box>
          <Button
            variant="contained"
            startIcon={<Download />}
            onClick={handleExport}
            disabled={generatedReports.length === 0}
          >
            전체 내보내기
          </Button>
        </Box>

        <Grid container spacing={2} sx={{ mb: 4 }}>
          {reports.map((report) => {
            const isGenerated = generatedReports.includes(report.id)
            const canGenerate = checkDependencies(report)

            return (
              <Grid item xs={12} sm={6} md={4} key={report.id}>
                <Card
                  variant="outlined"
                  sx={{
                    height: '100%',
                    borderColor: isGenerated
                      ? 'success.main'
                      : canGenerate
                      ? 'primary.main'
                      : 'divider',
                    opacity: canGenerate ? 1 : 0.6,
                  }}
                >
                  <CardContent>
                    <Box
                      sx={{
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center',
                        mb: 1,
                      }}
                    >
                      <Typography variant="h6">{report.name}</Typography>
                      {isGenerated && <CheckCircle color="success" />}
                    </Box>

                    <Typography
                      variant="body2"
                      color="text.secondary"
                      sx={{ mb: 2 }}
                    >
                      {report.description}
                    </Typography>

                    {!canGenerate && (
                      <Chip
                        label="데이터 준비 필요"
                        color="warning"
                        size="small"
                        sx={{ mb: 2 }}
                      />
                    )}

                    <Button
                      fullWidth
                      variant={isGenerated ? 'outlined' : 'contained'}
                      startIcon={<Description />}
                      onClick={() => handleGenerateReport(report)}
                      disabled={!canGenerate}
                    >
                      {isGenerated ? '재생성' : '리포트 생성'}
                    </Button>
                  </CardContent>
                </Card>
              </Grid>
            )
          })}
        </Grid>

        {generatedReports.length > 0 && (
          <TableContainer component={Paper} variant="outlined">
            <Table>
              <TableHead>
                <TableRow>
                  <TableCell>리포트명</TableCell>
                  <TableCell>상태</TableCell>
                  <TableCell>생성일시</TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {generatedReports.map((reportId) => {
                  const report = reports.find((r) => r.id === reportId)
                  return (
                    <TableRow key={reportId}>
                      <TableCell>{report?.name}</TableCell>
                      <TableCell>
                        <Chip label="생성 완료" color="success" size="small" />
                      </TableCell>
                      <TableCell>{new Date().toLocaleString('ko-KR')}</TableCell>
                    </TableRow>
                  )
                })}
              </TableBody>
            </Table>
          </TableContainer>
        )}

        <Box sx={{ display: 'flex', justifyContent: 'flex-end', mt: 3 }}>
          <Button variant="outlined" onClick={onBack}>
            이전
          </Button>
        </Box>
      </CardContent>
    </Card>
  )
}

export default ReportGenerationStep

