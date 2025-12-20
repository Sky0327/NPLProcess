// API 서비스 - Mock 데이터 반환

import {
  mockRegistryData,
  mockLandPriceData,
  mockAuctionData,
  mockKBData,
  mockRealEstatePriceData,
  mockInfocareData,
} from '../data/mockData'

// 시뮬레이션을 위한 딜레이 함수
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms))

export const apiService = {
  // 등기 조회
  async getRegistryInfo(inputData) {
    await delay(1500)
    return {
      success: true,
      data: mockRegistryData,
    }
  },

  // 공시지가 조회
  async getLandPrice(inputData) {
    await delay(1500)
    return {
      success: true,
      data: mockLandPriceData,
    }
  },

  // 법원 경매 조회
  async getCourtAuction(inputData) {
    await delay(1500)
    return {
      success: true,
      data: mockAuctionData,
    }
  },

  // KB 시세 조회
  async getKBPrice(depData) {
    await delay(2000)
    // depData는 { registry: [...], landPrice: [...] } 형태의 객체
    // 실제 구현에서는 depData를 활용하지만, mock에서는 무시
    return {
      success: true,
      data: mockKBData,
    }
  },

  // 실거래가 조회 (국토)
  async getRealEstatePrice(inputData) {
    await delay(1500)
    return {
      success: true,
      data: mockRealEstatePriceData,
    }
  },

  // 실거래가 조회 (밸류맵)
  async getValuemapPrice(inputData) {
    await delay(1500)
    return {
      success: true,
      data: mockRealEstatePriceData,
    }
  },

  // 거리 계산 (국토)
  async calculateDistance(depData) {
    await delay(1000)
    // depData는 { realEstatePrice: [...] } 형태의 객체
    const priceData = depData?.realEstatePrice || mockRealEstatePriceData || []
    const dataArray = Array.isArray(priceData) ? priceData : []
    return {
      success: true,
      data: dataArray.map((item) => ({
        ...item,
        distance: `${Math.floor(Math.random() * 1000)}m`,
      })),
    }
  },

  // 거리 계산 (밸류맵)
  async calculateDistanceValuemap(depData) {
    await delay(1000)
    // depData는 { valuemapPrice: [...] } 형태의 객체
    const priceData = depData?.valuemapPrice || mockRealEstatePriceData || []
    const dataArray = Array.isArray(priceData) ? priceData : []
    return {
      success: true,
      data: dataArray.map((item) => ({
        ...item,
        distance: `${Math.floor(Math.random() * 1000)}m`,
      })),
    }
  },

  // 인포케어 통계
  async getInfocareStats(inputData) {
    await delay(1500)
    return {
      success: true,
      data: mockInfocareData,
    }
  },

  // 인포케어 통합
  async getInfocareIntegrated(inputData) {
    await delay(1500)
    return {
      success: true,
      data: mockInfocareData,
    }
  },

  // 인포케어 사례
  async getInfocareCases(depData) {
    await delay(1500)
    // depData는 { infocareIntegrated: [...] } 형태의 객체
    // 실제 구현에서는 depData를 활용하지만, mock에서는 무시
    return {
      success: true,
      data: mockInfocareData,
    }
  },
}

