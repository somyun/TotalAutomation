#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
통합백그라운드.py
- [목적] Chrome 브라우저 실행 및 로그인 수행 후 '브라우저 유지(Detach)' 상태로 종료.
- [통신] 표준 입력(Stdin)으로 자격 증명 수신, 표준 출력(StdOut)으로 상태 전달.
- [구조] 메인 스레드(브라우저 제어) + 입력 스레드(Stdin 감시) 병렬 동작
- [로그] stdout: 사용자 상태 표시용 / file: 디버깅용 (debug_bg.txt)
"""

import sys
import json
import time
import threading
import subprocess
import traceback
import logging
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException

# 1. 표준 입출력 인코딩 (Windows 호환)
if sys.stdout:
    sys.stdout.reconfigure(encoding='cp949')
if sys.stdin:
    sys.stdin.reconfigure(encoding='utf-8')

# 2. 파일 로깅 설정 (디버그용)
logging.basicConfig(
    filename='debug_bg.txt',
    level=logging.DEBUG,
    format='[%(asctime)s] [%(levelname)s] (%(threadName)s) %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    encoding='utf-8',
    filemode='a' # Append mode
)

# 전역 상태
g_driver = None
g_lock = threading.Lock()
g_credentials = {} # {id, pw, pw2}

# 동기화 이벤트
evt_page_ready = threading.Event()
evt_auth_received = threading.Event()
evt_shutdown = threading.Event()

# ==============================================================================
# 로깅 헬퍼 함수
# ==============================================================================
SENSITIVE_KEYS = {"id", "pw", "pw2"}


def _sanitize_ipc_for_log(text: str) -> str:
    """
    IPC 원문에서 민감 필드(id/pw/pw2)를 길이 정보만 남기도록 마스킹.
    JSON 한 줄이 아니라 조각("id":"...", 등)만 들어와도 동작하도록 정규식 기반으로 처리.
    """

    def repl(m: re.Match) -> str:
        key = m.group(1)
        value = m.group(2)
        length = len(value) if value is not None else 0
        return f'"{key}":"len={length}"'

    try:
        masked = re.sub(
            r'"(id|pw|pw2)"\s*:\s*"([^"]*)"',
            repl,
            text,
        )
        # 전체 길이도 힌트로 남기되, 실제 내용은 포함하지 않음
        return f"{masked} (raw_len={len(text)})"
    except Exception:
        # 혹시 마스킹 중 문제가 나면, 내용은 버리고 길이만 남김
        return f"<masked_ipc len={len(text)}>"


def log_debug(msg, level="INFO"):
    """파일에만 기록 (사용자에게 보이지 않음)"""
    if level == "ERROR":
        logging.error(msg)
    elif level == "WARN":
        logging.warning(msg)
    else:
        logging.info(msg)

def log_user(msg, tag="INFO"):
    """AHK(UI)로 전송 (파일에도 기록됨)"""
    # 1. 파일 기록
    log_debug(f"[UI_SEND] ({tag}) {msg}")

    # 2. Stdout 전송 (Flush 필수)
    try:
        if tag == "STATE":
            print(f"STATE:{msg}", flush=True)
        else:
            # tag가 INFO 등일 때는 그냥 메시지만 출력하여 UI 타이틀바에 표시
            print(msg, flush=True)
    except OSError:
        pass

# ==============================================================================
# 입력 리스너 (IPC)
# ==============================================================================
def input_listener():
    """Stdin 감시 데몬 스레드"""
    log_debug("Input listener thread started")
    
    while not evt_shutdown.is_set():
        try:
            line = sys.stdin.readline()
            if not line:
                log_debug("Stdin closed (EOF)", "WARN")
                break 
            
            line = line.strip()
            if not line:
                continue

            # 로깅은 파일에만 (보안/성능) / 민감정보는 길이만 남기고 마스킹
            safe_line = _sanitize_ipc_for_log(line)
            log_debug(f"Received IPC: {safe_line[:80]}...", "INFO") 

            try:
                cmd = json.loads(line)
                cmd_type = cmd.get("type")

                if cmd_type == "login":
                    log_debug("Credential command received")
                    with g_lock:
                        g_credentials["id"] = cmd.get("id")
                        g_credentials["pw"] = cmd.get("pw")
                        g_credentials["pw2"] = cmd.get("pw2")
                    evt_auth_received.set()
                    log_user("AUTH_RECEIVED", "STATE")
                    log_user("인증 정보 수신 완료")
                
                elif cmd_type == "exit":
                    log_debug("Exit command received")
                    evt_shutdown.set()
                    
            except json.JSONDecodeError:
                # JSON 파싱 실패 시에도 원문 전체를 남기지 않고, 마스킹/길이 정보만 남김
                safe_line = _sanitize_ipc_for_log(line)
                log_debug(f"Invalid JSON: {safe_line}", "WARN")

        except Exception as e:
            log_debug(f"Input listener error: {e}", "ERROR")
            break
    
    log_debug("Input listener thread stopped")

# ==============================================================================
# 브라우저 자동화 클래스
# ==============================================================================
class ChromeAutomation:
    def __init__(self, port=9222, headless=True):
        self.port = port
        self.headless = headless
        self.driver = None
        self._setup_driver()

    def _setup_driver(self):
        global g_driver
        chrome_options = Options()
        
        # Detach: 프로세스 유지
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_argument(f"--remote-debugging-port={self.port}")

        if self.headless:
            chrome_options.add_argument("--headless=new") 
        
        # 탐지 회피 및 성능 옵션
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--ignore-certificate-errors")
        chrome_options.add_argument("--disable-infobars")
        chrome_options.add_argument("--log-level=3") 
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            g_driver = self.driver
            
            log_debug(f"Chrome initialized on port {self.port}")
            log_user("BG_START", "STATE")

        except WebDriverException as e:
            log_debug(f"Port {self.port} busy or init failed: {e}", "WARN")
            if self._kill_existing_browser():
                try:
                    time.sleep(1)
                    self.driver = webdriver.Chrome(options=chrome_options)
                    self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
                    log_debug("Chrome restarted successfully")
                    log_user("BG_START", "STATE")
                    return
                except Exception as e2:
                    log_debug(f"Retry failed: {e2}", "ERROR")
            
            log_debug(f"Fatal init error: {e}", "ERROR")
            sys.exit(1)

    def _kill_existing_browser(self):
        try:
            cmd = f'wmic process where "CommandLine like \'%--remote-debugging-port={self.port}%\'" call terminate'
            subprocess.run(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except Exception:
            return False

    def navigate(self, url):
        try:
            self.driver.get(url)
            return True
        except Exception as e:
            log_debug(f"Navigate error ({url}): {e}", "ERROR")
            return False

class LoginManager:
    def __init__(self, headless=True):
        self.portal_url = "https://btcep.humetro.busan.kr/portal"
        self.ep_url = "https://niw.humetro.busan.kr/erpep.jsp"
        self.chrome = None
        self.headless = headless

    def run(self):
        try:
            # 1. 브라우저 시작
            self.chrome = ChromeAutomation(headless=self.headless)
            
            # 2. 포털 이동
            log_user("포털 이동 중...")
            if self.chrome.navigate(self.portal_url):
                log_debug("Page loaded, waiting for auth event")
                log_user("PAGE_READY", "STATE")
                evt_page_ready.set()
            else:
                raise Exception("Portal navigation failed")

            # 3. 로그인 정보 대기
            # (Polling 로그는 파일에도 남기지 않음, 이벤트만 기다림)
            if not evt_auth_received.wait(timeout=300):
                log_debug("Auth wait timeout (300s)", "ERROR")
                return

            if evt_shutdown.is_set():
                log_debug("Shutdown signal received before login")
                return

            # 4. 로그인 수행
            log_user("LOGIN_START", "STATE")
            log_user("인증 수행 중...")
            if self._perform_login():
                if self._move_to_ep():
                    log_user("LOGIN_OK", "STATE")
                    log_user("로그인 및 EP 이동 완료")
                    
                    time.sleep(2) # AHK가 메시지 처리할 시간 여유
                    return 
            
            # 실패
            log_user("LOGIN_FAIL", "STATE")
            self._handle_failure()

        except Exception as e:
            log_debug(f"Main run loop error: {e}", "ERROR")
            log_debug(traceback.format_exc(), "ERROR")
            self._handle_failure()

    def _perform_login(self):
        try:
            uid, upw, ucerti = "", "", ""
            with g_lock:
                uid = g_credentials.get("id")
                upw = g_credentials.get("pw")
                ucerti = g_credentials.get("pw2")

            driver = self.chrome.driver
            wait = WebDriverWait(driver, 10)
            
            # 로그인 시도 시에도 실제 아이디 값 대신 길이만 남겨서 추적
            uid_len = len(uid) if uid else 0
            log_debug(f"Attempting login for user (id_len={uid_len})")

            # 아이디/비번
            wait.until(EC.presence_of_element_located((By.ID, "userId")))
            driver.execute_script(f"document.querySelector('#userId').value = '{uid}'")
            driver.execute_script(f"document.querySelector('#password').value = '{upw}'")
            driver.execute_script("document.querySelector('.btn_login').click()")
            
            # 2차 인증
            log_user("2차 인증 중...")
            wait.until(EC.presence_of_element_located((By.ID, "certi_num")))
            driver.execute_script(f"document.querySelector('#certi_num').value = '{ucerti}'")
            driver.execute_script("login()")
            
            # 완료 대기 (로그아웃 버튼)
            log_user("인증 확인 중...")
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'로그아웃')] | //a[contains(@href, 'logout')]"))
            )
            log_debug("Login verification success")
            return True
            
        except Exception as e:
            log_debug(f"Login logic failed: {e}", "ERROR")
            return False

    def _move_to_ep(self):
        try:
            log_user("EP 시스템 이동 중...")
            driver = self.chrome.driver
            
            # 새 탭
            driver.execute_script("window.open('about:blank','_blank');")
            driver.switch_to.window(driver.window_handles[-1])
            driver.get(self.ep_url)
            
            # EP 대기
            log_user("EP 로딩 중...")
            WebDriverWait(driver, 40).until(
                EC.presence_of_element_located((By.ID, "goLogOut"))
            )
            log_debug("EP load success")
            return True
        except Exception as e:
            log_debug(f"EP move failed: {e}", "ERROR")
            return False

    def _handle_failure(self):
        try:
            if self.chrome and self.chrome.driver:
                self.chrome.driver.quit()
        except:
            pass
        sys.exit(1)

def main():
    log_debug("=== Process Started ===")
    
    input_thread = threading.Thread(target=input_listener, daemon=True, name="InputThread")
    input_thread.start()
    
    manager = LoginManager(headless=True)
    manager.run()
    
    log_debug("=== Process Finished ===")

if __name__ == "__main__":
    main()