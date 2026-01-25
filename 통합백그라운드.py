#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
통합백그라운드.py (구 jsession.py)
- [목적] Chrome 브라우저 실행 및 로그인 수행 후 '브라우저 유지(Detach)' 상태로 종료.
- [통신] 표준 출력(StdOut)을 통해 진행 상황을 AHK로 전달.
"""

import sys
import json
import argparse
import time
import subprocess
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, WebDriverException

# 표준 출력 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

def log(msg):
    """AHK가 읽을 수 있도록 표준 출력으로 메시지 전송 (Flush 필수)"""
    print(msg, flush=True)

class ChromeAutomation:
    def __init__(self, port=9222, headless=True):
        self.port = port
        self.headless = headless
        self.driver = None
        self._setup_driver()

    def _setup_driver(self):
        chrome_options = Options()
        
        # [핵심] Detach 옵션: 스크립트 종료 후에도 브라우저 프로세스 유지
        chrome_options.add_experimental_option("detach", True)
        
        # 디버깅 포트 설정
        chrome_options.add_argument(f"--remote-debugging-port={self.port}")

        if self.headless:
            chrome_options.add_argument("--headless=new") # 최신 헤드리스 모드 권장
        
        # 봇 탐지 회피
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # 성능 및 안정성
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--ignore-certificate-errors")
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            # WebDriver 속성 숨기기
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            log(f"Browser Started on Port {self.port}")
        except WebDriverException as e:
            # 이미 포트가 사용 중일 수 있음 (기존 브라우저가 안 닫혔거나)
            log(f"Error: Chrome init failed. {e}")
            sys.exit(1)

    def navigate(self, url):
        try:
            self.driver.get(url)
            return True
        except Exception as e:
            log(f"Error: Navigate failed - {e}")
            return False

class LoginManager:
    def __init__(self, user_id, user_pw, user_pw2, headless=True):
        self.id = user_id
        self.pw = user_pw
        self.certi = user_pw2
        self.portal_url = "https://btcep.humetro.busan.kr/portal"
        self.ep_url = "https://niw.humetro.busan.kr/erpep.jsp"
        
        # 브라우저 실행
        self.chrome = ChromeAutomation(headless=headless)

    def run(self):
        if self._login_portal():
            self._move_to_ep()
            log("Ready") # AHK에 제어권 인계 신호
            # 스크립트는 여기서 종료되지만, detach=True로 브라우저는 살아있음
        else:
            log("Fail")

    def _login_portal(self):
        try:
            log("포털 이동 중...")
            if not self.chrome.navigate(self.portal_url): return False
            
            # 1차 로그인
            log("1차 인증 중...")
            WebDriverWait(self.chrome.driver, 15).until(EC.presence_of_element_located((By.ID, "userId")))
            self.chrome.driver.execute_script(f"document.querySelector('#userId').value = '{self.id}'")
            self.chrome.driver.execute_script(f"document.querySelector('#password').value = '{self.pw}'")
            self.chrome.driver.execute_script("document.querySelector('.btn_login').click()")
            
            # 2차 인증
            log("2차 인증 중...")
            WebDriverWait(self.chrome.driver, 15).until(EC.presence_of_element_located((By.ID, "certi_num")))
            self.chrome.driver.execute_script(f"document.querySelector('#certi_num').value = '{self.certi}'")
            self.chrome.driver.execute_script("login()")
            
            # 완료 대기
            log("인증 확인 중...")
            WebDriverWait(self.chrome.driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'로그아웃')] | //a[contains(@href, 'logout')]"))
            )
            return True
        except Exception as e:
            log(f"Error: Login failed - {e}")
            return False

    def _move_to_ep(self):
        try:
            log("EP 시스템 이동 중...")
            # 탭 정리 및 EP 이동 로직 (기존 유지)
            # 새 탭 열기 (기존 탭 간섭 방지)
            self.chrome.driver.execute_script("window.open('about:blank','_blank');")
            self.chrome.driver.switch_to.window(self.chrome.driver.window_handles[-1])
            
            self.chrome.navigate(self.ep_url)
            
            # EP 로딩 대기
            log("EP 로딩 중...")
            WebDriverWait(self.chrome.driver, 40).until(
                EC.presence_of_element_located((By.ID, "goLogOut"))
            )
        except Exception as e:
            log(f"Error: EP Move failed - {e}")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--id', required=True)
    parser.add_argument('--pw', required=True)
    parser.add_argument('--pw2', required=True)
    parser.add_argument('--headless', action='store_true', help='Run in headless mode')
    
    args = parser.parse_args()
    
    # Headless 모드는 기본값이 True이지만 인자로 제어 가능하게 설정
    is_headless = True
    # 디버깅용으로 --show-browser 같은게 필요할 수 있으나 일단 고정
    
    manager = LoginManager(args.id, args.pw, args.pw2, headless=is_headless)
    manager.run()

if __name__ == "__main__":
    main()