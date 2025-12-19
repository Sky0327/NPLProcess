import React from 'react'
import { useNavigate } from 'react-router-dom'
import {
  Box,
  Typography,
  Card,
  CardContent,
  CardActions,
  Button,
  Grid,
  Paper,
} from '@mui/material'
import {
  Assessment,
  Description,
  Settings,
  CloudUpload,
} from '@mui/icons-material'

const Home = () => {
  const navigate = useNavigate()

  const features = [
    {
      title: 'Smart_NPL1',
      description: '데이터 수집 및 처리',
      icon: <Assessment sx={{ fontSize: 48 }} />,
      path: '/smart-npl1',
      color: 'primary',
    },
    {
      title: 'Smart_NPL2',
      description: '리포트 생성 및 내보내기',
      icon: <Description sx={{ fontSize: 48 }} />,
      path: '/smart-npl2',
      color: 'secondary',
    },
  ]

  return (
    <Box>
      <Box sx={{ mb: 4, textAlign: 'center' }}>
        <Typography variant="h3" component="h1" gutterBottom sx={{ fontWeight: 700 }}>
          Samil NPL 평가 시스템
        </Typography>
        <Typography variant="h6" color="text.secondary" sx={{ mt: 2 }}>
          부실채권 평가를 위한 통합 데이터 수집 및 리포트 생성 시스템
        </Typography>
      </Box>

      <Grid container spacing={3} sx={{ mb: 4 }}>
        {features.map((feature) => (
          <Grid item xs={12} md={6} key={feature.title}>
            <Card
              sx={{
                height: '100%',
                display: 'flex',
                flexDirection: 'column',
                cursor: 'pointer',
                transition: 'transform 0.2s, box-shadow 0.2s',
                '&:hover': {
                  transform: 'translateY(-4px)',
                  boxShadow: 6,
                },
              }}
              onClick={() => navigate(feature.path)}
            >
              <CardContent sx={{ flexGrow: 1, textAlign: 'center', pt: 4 }}>
                <Box sx={{ color: `${feature.color}.main`, mb: 2 }}>
                  {feature.icon}
                </Box>
                <Typography variant="h5" component="h2" gutterBottom>
                  {feature.title}
                </Typography>
                <Typography variant="body1" color="text.secondary">
                  {feature.description}
                </Typography>
              </CardContent>
              <CardActions sx={{ justifyContent: 'center', pb: 3 }}>
                <Button
                  variant="contained"
                  color={feature.color}
                  size="large"
                  onClick={() => navigate(feature.path)}
                >
                  시작하기
                </Button>
              </CardActions>
            </Card>
          </Grid>
        ))}
      </Grid>

      <Paper sx={{ p: 3, backgroundColor: 'background.paper' }}>
        <Typography variant="h6" gutterBottom sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
          <Settings />
          시스템 개요
        </Typography>
        <Typography variant="body2" color="text.secondary" sx={{ mt: 2 }}>
          본 시스템은 Excel VBA 기반 NPL 평가 시스템을 웹 기반으로 마이그레이션한 것입니다.
          등기 정보, 공시지가, 경매 정보, 부동산 가격 정보 등을 수집하고 분석하여
          체계적인 평가 리포트를 생성합니다.
        </Typography>
      </Paper>
    </Box>
  )
}

export default Home

