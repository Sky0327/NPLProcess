import React, { useState } from 'react'
import {
  Box,
  Typography,
  Card,
  CardContent,
  Grid,
  Button,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
} from '@mui/material'
import {
  Description,
  PictureAsPdf,
  FileDownload,
  Refresh,
} from '@mui/icons-material'

const SmartNPL2 = () => {
  const [reports, setReports] = useState([
    { id: 'report-0', name: '[0] 물건지', generated: false },
    { id: 'report-1', name: '[1] 채권현황', generated: false },
  ])

  const handleGenerateReport = (reportId) => {
    setReports(
      reports.map((r) =>
        r.id === reportId ? { ...r, generated: true } : r
      )
    )
    alert('리포트가 생성되었습니다.')
  }

  const handleConvertPDF = () => {
    alert('PDF 파일명이 변환되었습니다.')
  }

  const handleExportXLSX = () => {
    alert('Excel 파일로 내보내기가 완료되었습니다.')
  }

  const handleReset = () => {
    if (window.confirm('파일을 초기화하시겠습니까?')) {
      setReports(reports.map((r) => ({ ...r, generated: false })))
      alert('파일이 초기화되었습니다.')
    }
  }

  return (
    <Box>
      <Typography variant="h4" component="h1" gutterBottom sx={{ mb: 4, fontWeight: 700 }}>
        Smart_NPL2 - 리포트 생성 및 내보내기
      </Typography>

      <Grid container spacing={3}>
        <Grid item xs={12} md={6}>
          <Card>
            <CardContent>
              <Typography variant="h6" gutterBottom>
                최종 리포트 생성
              </Typography>
              <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2, mt: 2 }}>
                {reports.map((report) => (
                  <Button
                    key={report.id}
                    variant={report.generated ? 'outlined' : 'contained'}
                    startIcon={<Description />}
                    onClick={() => handleGenerateReport(report.id)}
                    fullWidth
                  >
                    {report.name} {report.generated ? '(생성됨)' : ''}
                  </Button>
                ))}
              </Box>
            </CardContent>
          </Card>
        </Grid>

        <Grid item xs={12} md={6}>
          <Card>
            <CardContent>
              <Typography variant="h6" gutterBottom>
                파일 관리
              </Typography>
              <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2, mt: 2 }}>
                <Button
                  variant="contained"
                  color="secondary"
                  startIcon={<PictureAsPdf />}
                  onClick={handleConvertPDF}
                  fullWidth
                >
                  등본PDF 파일명 변환
                </Button>
                <Button
                  variant="contained"
                  color="primary"
                  startIcon={<FileDownload />}
                  onClick={handleExportXLSX}
                  fullWidth
                  disabled={!reports.some((r) => r.generated)}
                >
                  xlsx로 내보내기
                </Button>
                <Button
                  variant="outlined"
                  startIcon={<Refresh />}
                  onClick={handleReset}
                  fullWidth
                >
                  파일 초기화
                </Button>
              </Box>
            </CardContent>
          </Card>
        </Grid>

        <Grid item xs={12}>
          <Card>
            <CardContent>
              <Typography variant="h6" gutterBottom>
                생성된 리포트 목록
              </Typography>
              <TableContainer component={Paper} variant="outlined" sx={{ mt: 2 }}>
                <Table>
                  <TableHead>
                    <TableRow>
                      <TableCell>리포트명</TableCell>
                      <TableCell>상태</TableCell>
                      <TableCell>생성일시</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {reports
                      .filter((r) => r.generated)
                      .map((report) => (
                        <TableRow key={report.id}>
                          <TableCell>{report.name}</TableCell>
                          <TableCell>생성 완료</TableCell>
                          <TableCell>{new Date().toLocaleString('ko-KR')}</TableCell>
                        </TableRow>
                      ))}
                    {reports.filter((r) => r.generated).length === 0 && (
                      <TableRow>
                        <TableCell colSpan={3} align="center">
                          생성된 리포트가 없습니다.
                        </TableCell>
                      </TableRow>
                    )}
                  </TableBody>
                </Table>
              </TableContainer>
            </CardContent>
          </Card>
        </Grid>
      </Grid>
    </Box>
  )
}

export default SmartNPL2

