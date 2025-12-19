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
  async getKBPrice(registryData, landPriceData) {
    await delay(2000)
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
  async calculateDistance(priceData) {
    await delay(1000)
    return {
      success: true,
      data: priceData.map((item) => ({
        ...item,
        distance: `${Math.floor(Math.random() * 1000)}m`,
      })),
    }
  },

  // 거리 계산 (밸류맵)
  async calculateDistanceValuemap(priceData) {
    await delay(1000)
    return {
      success: true,
      data: priceData.map((item) => ({
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
  async getInfocareCases(integratedData) {
    await delay(1500)
    return {
      success: true,
      data: mockInfocareData,
    }
  },
}

