#!/usr/bin/env python3
# ==============================================================================
# hdmi_gen_test.py
# Matrox SQA - HDMI Generator Test Script
# Replaces HdmiGenWithWebCam.cs / HostRhesus with a standalone Python script
# Controls the Digilent Nexys Video FPGA board over serial COM port,
# cycles through HDMI resolutions and defects, captures webcam images,
# runs ImageProcess.exe for colorbar recognition, logs pass/fail results.
#
# Usage:
#   python hdmi_gen_test.py [options]
#
# Dependencies:
#   pip install pyserial
# ==============================================================================

import argparse
import os
import random
import subprocess
import sys
import time
import datetime
import serial  # pip install pyserial
import requests
import urllib3

# ==============================================================================
# Paths to external tools
# ==============================================================================
COMMAND_CAM_EXE   = r"C:\HostRhesus\CommandCam.exe"
IMAGE_PROCESS_EXE = r"C:\cprogram\ImageProcess\ImageProcess\bin\Debug\net8.0\ImageProcess.exe"
IMAGE_BMP_PATH    = r"C:\HostRhesus\image.bmp"           # CommandCam writes here
LOG_DIR           = os.path.join(os.environ.get("USERPROFILE", "C:\\Users\\rabdulla"), "Desktop")
DEFAULT_LOG_FILE  = os.path.join(LOG_DIR, "HDMILog.log")

# ==============================================================================
# Avio2 API configuration
# ==============================================================================
AVIO2_TX_IP          = "192.168.189.174"
AVIO2_AUTH_URL       = f"https://{AVIO2_TX_IP}/auth/v1/users/login"
AVIO2_STREAM_URL     = f"https://{AVIO2_TX_IP}/app/v2/device/status/streams/video/0"
AVIO2_HEALTH_URL     = f"https://{AVIO2_TX_IP}/mgmt/v1/healthstatus/ishealthy"
AVIO2_USERNAME       = "Tester"
AVIO2_PASSWORD       = "Matrox1234!"
STREAM_STATE_ACTIVE  = 2    # StreamSignalState: 1=Disabled, 2=Active, 3=Inactive, 4=Blanked
RELOCK_WAIT_S        = 15   # seconds to wait after re-lock before resuming

# ==============================================================================
# User-configurable parameters (all overridable via CLI args)
# ==============================================================================
DEFAULT_COM_PORT              = "COM8"
DEFAULT_CAM_INDEX             = 1
DEFAULT_STOP_AFTER_ERROR      = 0       # 0 = never stop on error
DEFAULT_JITTER_AMPLITUDE_MBIT = 300     # mTbit  (0.300 Tbit = 300)
DEFAULT_TRANSITIONAL_JITTER   = True
DEFAULT_JITTER_TIME_MS        = 500
DEFAULT_TRANSITIONAL_INTERPAIR= True
DEFAULT_INTERPAIR_DELAY_MS    = 500
DEFAULT_LOOP_DELAY_MS         = 15000  # ms to wait before webcam snapshot
DEFAULT_SOURCE_MASK           = 0x7E00 # enabled resolutions bitmask
DEFAULT_SOURCE_BLACK_MASK     = 0x0000
DEFAULT_SOURCE_ALGO           = "INCREMENTAL"
DEFAULT_SOURCE_START          = 11     # HDMI_1920x1080P60_148500_RGB_8
DEFAULT_DEFECT_MASK           = 0x07FC0 # all defects enabled
DEFAULT_DEFECT_IMAGE_MASK     = 0xFFFF
DEFAULT_DEFECT_ALGO           = "INCREMENTAL"
DEFAULT_DEFECT_START          = 0      # CLEAN
DEFAULT_PING_IP               = ""     # empty = no ping
DEFAULT_PING_ENABLE           = False
DEFAULT_IMAGE_THRESHOLD       = 60     # % colorbar recognition to call PASS

# ==============================================================================
# Enums as integer constants (match ESourceRes / EDefectType in C# script)
# ==============================================================================
SOURCE_NAMES = [
    "DVI_720x480P60_27027_RGB_8",
    "HDMI_640x480P60_25200_RGB_8",
    "HDMI_720x480P60_27027_RGB_8",
    "HDMI_720x576P50_27000_RGB_8",
    "HDMI_3840x2160P60_594000_YUV420_8",
    "HDMI_1280x720P60_74250_RGB_8",
    "HDMI_1280x720P50_92813_RGB_10",
    "HDMI_1920x1080I50_74250_RGB_8",
    "HDMI_1920x1080I60_74250_RGB_8",
    "HDMI_1920x1080P30_74250_RGB_8",
    "HDMI_1920x1080P50_148500_RGB_8",
    "HDMI_1920x1080P60_148500_RGB_8",
    "HDMI_1920x1080P60_185625_RGB_10",
    "HDMI_1920x1080P60_222750_RGB_12",
    "HDMI_3840x2160P24_297000_RGB_8",
]
SOURCE_COUNT = len(SOURCE_NAMES)

DEFECT_NAMES = [
    "CLEAN",
    "CLOCK_JITTER",
    "SYMBOL_SKEW_DELAY",
    "SEQUENTIAL_DATA",
    "TMDS_OUTPUT_ENABLE",
    "TMDS_SCRAMBLING",
    "RANDOM_5V_DISCONNECT",
    "TMDS_SWAP",
    "ALL_DRESS",
    "BLACK",
]
DEFECT_COUNT = len(DEFECT_NAMES)

# Source timing info: (ResX, ResY, RefreshRate, IsInterlaced, PixelClock_kHz, DC_bits, Y1Y0, IsHdmi)
# Y1Y0: 0=RGB, 1=422, 2=444, 3=420
SOURCE_INFO = [
    (720,  480,  60, False,  27027,  8, 0, False),  # 0  DVI_720x480P60
    (640,  480,  60, False,  25203,  8, 0, True ),  # 1  HDMI_640x480P60
    (720,  480,  60, False,  27027,  8, 0, True ),  # 2  HDMI_720x480P60
    (720,  576,  50, False,  27000,  8, 0, True ),  # 3  HDMI_720x576P50
    (3840, 2160, 60, False, 594000,  8, 3, True ),  # 4  HDMI_3840x2160P60_YUV420
    (1280, 720,  60, False,  74250,  8, 0, True ),  # 5  HDMI_1280x720P60
    (1280, 720,  50, False,  74250, 10, 0, True ),  # 6  HDMI_1280x720P50_10bit
    (1920, 1080, 25, True,   74250,  8, 0, True ),  # 7  HDMI_1920x1080I50
    (1920, 1080, 30, True,   74250,  8, 0, True ),  # 8  HDMI_1920x1080I60
    (1920, 1080, 30, False,  74250,  8, 0, True ),  # 9  HDMI_1920x1080P30
    (1920, 1080, 50, False, 148500,  8, 0, True ),  # 10 HDMI_1920x1080P50
    (1920, 1080, 60, False, 148500,  8, 0, True ),  # 11 HDMI_1920x1080P60
    (1920, 1080, 60, False, 148500, 10, 0, True ),  # 12 HDMI_1920x1080P60_10bit
    (1920, 1080, 60, False, 148333, 12, 0, True ),  # 13 HDMI_1920x1080P60_12bit
    (3840, 2160, 24, False, 297484,  8, 0, True ),  # 14 HDMI_3840x2160P24
]

# COM commands: (clock_cmd, res_cmd)  - matches SetHdmiSource() in C# script
SOURCE_COMMANDS = [
    ("c", "l"),  # 0  DVI_720x480P60
    ("a", "m"),  # 1  HDMI_640x480P60
    ("c", "l"),  # 2  HDMI_720x480P60  (same as DVI per C# script bug - kept faithful)
    ("b", "o"),  # 3  HDMI_720x576P50
    ("j", "p"),  # 4  HDMI_3840x2160P60_YUV420
    ("d", "q"),  # 5  HDMI_1280x720P60
    ("e", "r"),  # 6  HDMI_1280x720P50_10bit
    ("d", "s"),  # 7  HDMI_1920x1080I50
    ("d", "t"),  # 8  HDMI_1920x1080I60
    ("d", "u"),  # 9  HDMI_1920x1080P30
    ("g", "v"),  # 10 HDMI_1920x1080P50
    ("g", "w"),  # 11 HDMI_1920x1080P60
    ("h", "x"),  # 12 HDMI_1920x1080P60_10bit
    ("i", "y"),  # 13 HDMI_1920x1080P60_12bit
    ("j", "z"),  # 14 HDMI_3840x2160P24
]


# ==============================================================================
# Serial COM port wrapper
# ==============================================================================
class HdmiGenSerial:
    def __init__(self, port, baud=115200):
        self.ser = serial.Serial(
            port=port,
            baudrate=baud,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=1,
            xonxoff=False,
            rtscts=False,
        )
        time.sleep(0.1)

    def send(self, cmd):
        """Send ASCII command string to FPGA board."""
        if isinstance(cmd, str):
            cmd = cmd.encode("ascii")
        self.ser.write(cmd)
        time.sleep(0.05)  # small settle time between commands

    def close(self):
        if self.ser.is_open:
            self.ser.close()


# ==============================================================================
# Logging
# ==============================================================================
class Logger:
    def __init__(self, log_file):
        self.log_file = log_file
        self.console_width = 200  # wide for iteration lines like original

    def log(self, msg, newline=True):
        """Print to console and append to log file."""
        end = "\n" if newline else ""
        print(msg, end=end, flush=True)
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(msg + end)

    def log_inline(self, msg):
        """Print without newline (for building up a single iteration line)."""
        self.log(msg, newline=False)


# ==============================================================================
# Source info string builder (matches ReadSourceInfo() in C# script)
# ==============================================================================
def source_info_str(src_idx):
    rx, ry, rate, interlaced, pclk, dc, y1y0, is_hdmi = SOURCE_INFO[src_idx]
    scan = "i" if interlaced else "p"
    color = {0: "RGB", 1: "422", 2: "444", 3: "420"}.get(y1y0, "???")
    iface = "HDMI" if is_hdmi else " DVI"
    return f"{rx:4d} x {ry:4d}{scan} @ {rate:2d} Hz {pclk:6d} kHz {dc:2d} bits {color} {iface} "


# ==============================================================================
# FPGA control functions (mirror C# methods exactly)
# ==============================================================================
def init_registers(ser):
    """InitRegistersHdmiGen - reset all FPGA control registers to safe defaults."""
    ser.send("W300")  # No jitter
    ser.send("W400")  # Enable lane OEn (0 = enabled)
    ser.send("W501")  # Enable HDMI +5V
    ser.send("W600")  # Enable TMDS141 OEN
    ser.send("W7E4")  # Default lane swap: CLK D2 D1 D0
    ser.send("W800")  # Scrambling disabled
    ser.send("W900")  # No inter-pair delay D1/D0
    ser.send("WA00")  # No inter-pair delay CLK/D2


def disable_hdmi(ser):
    ser.send(",")     # Disable HDMI datapath
    ser.send("W40F")  # HI-Z all TMDS lanes
    ser.send("W500")  # Disable HDMI +5V
    ser.send("W601")  # Disable TMDS141 OEN


def enable_hdmi(ser):
    ser.send(".")     # Enable HDMI TMDS datapath
    ser.send("W400")  # Enable TMDS lanes (no HI-Z)
    ser.send("W501")  # Enable HDMI +5V
    ser.send("W600")  # Enable TMDS141 OEN


def set_hdmi_source(ser, src_idx):
    """SetHdmiSource - send clock + resolution + reset commands."""
    clk_cmd, res_cmd = SOURCE_COMMANDS[src_idx]
    ser.send(clk_cmd + res_cmd + "!")  # clock, resolution, reset


def set_clock_jitter(ser, src_idx, jitter_amplitude_mbit):
    """SetClockJitter - calculate and apply jitter register value."""
    _, _, _, _, pclk_khz, dc, _, _ = SOURCE_INFO[src_idx]

    tmds_khz = pclk_khz
    if dc == 10:
        tmds_khz = (tmds_khz * 125 + 50) // 100
    elif dc == 12:
        tmds_khz = (tmds_khz * 150 + 50) // 100

    # VCO frequency lookup (matches C# constants)
    vco_table = [
        (26000,   630000),
        (27010,   945000),
        (50000,  1081250),
        (80000,   742500),
        (100000,  928125),
        (160000,  742500),
        (200000,  928125),
        (250000, 1113750),
        (float("inf"), 1485000),
    ]
    vco_khz = 1485000
    for limit, vco in vco_table:
        if tmds_khz < limit:
            vco_khz = vco
            break

    if tmds_khz == 0 or vco_khz == 0:
        return

    jitter_num = (jitter_amplitude_mbit * 100000) / tmds_khz
    jitter_den = 1_000_000_000 / vco_khz / 56
    if jitter_den == 0:
        return

    amp = int(jitter_num / jitter_den + 0.5)
    amp = min(amp, 255)
    ser.send(f"W3{amp:02X}")


def inter_pair_delay_random(ser, max_delay, transitional, delay_ms):
    d0  = random.randint(0, max_delay)
    d1  = random.randint(0, max_delay)
    d2  = random.randint(0, max_delay)
    clk = random.randint(0, max_delay)
    reg9 = (d1 << 4) | d0
    rega = (clk << 4) | d2
    ser.send(f"W9{reg9:02X}")
    ser.send(f"WA{rega:02X}")
    if transitional:
        time.sleep(delay_ms / 1000.0)
        ser.send("W900")
        ser.send("WA00")


def sequential_data_random(ser):
    ser.send("W40F")   # Disable all lanes
    time.sleep(0.1)
    ser.send(".")       # Enable datapath
    time.sleep(0.1)
    max_ms = 100
    d0  = random.randint(0, max_ms)
    d1  = random.randint(0, max_ms)
    d2  = random.randint(1, max_ms)
    clk = random.randint(1, max_ms)
    quartet = 0x0F
    for t in range(max_ms + 1):
        if t == d0:
            quartet &= 0x0E
            ser.send(f"W40{quartet:01X}")
        if t == d1:
            quartet &= 0x0D
            ser.send(f"W40{quartet:01X}")
        if t == d2:
            quartet &= 0x0B
            ser.send(f"W40{quartet:01X}")
        if t == clk:
            quartet &= 0x07
            ser.send(f"W40{quartet:01X}")
        time.sleep(0.001)


def glitchy_output_enable_random(ser):
    ser.send(".")
    n = random.randint(2, 15)
    for _ in range(n):
        off_ms = random.randint(0, 250)
        on_ms  = random.randint(0, 250)
        ser.send("W601")          # Disable TMDS141 OEN
        time.sleep(off_ms / 1000.0)
        ser.send("W600")          # Enable TMDS141 OEN
        time.sleep(on_ms / 1000.0)


def scrambling_random(ser):
    ser.send(".")
    n = random.randint(2, 15)
    for _ in range(n):
        on_ms  = random.randint(0, 250)
        off_ms = random.randint(0, 250)
        ser.send("W80F")          # Enable scrambling all lanes
        time.sleep(on_ms / 1000.0)
        ser.send("W800")          # Disable scrambling
        time.sleep(off_ms / 1000.0)


def disconnect_5v_random(ser):
    ser.send(".")
    n = random.randint(2, 15)
    for _ in range(n):
        off_ms = random.randint(0, 250)
        on_ms  = random.randint(0, 250)
        ser.send("W500")          # Disable HDMI +5V
        time.sleep(0.250)         # fixed 250ms off (matches C# iMaxOffTime_ms)
        ser.send("W501")          # Enable HDMI +5V
        time.sleep(on_ms / 1000.0)


def swap_tmds_pair_random(ser):
    swap_cmds = ["W7E4", "W7E1", "W7D8", "W7D2", "W7C9", "W7C6"]
    ser.send(".")
    n = random.randint(2, 15)
    for _ in range(n):
        on_ms  = random.randint(0, 250)
        off_ms = random.randint(0, 250)
        cmd = random.choice(swap_cmds)
        ser.send(cmd)
        time.sleep(on_ms / 1000.0)
        ser.send("W7E4")          # Restore default
        time.sleep(off_ms / 1000.0)


# ==============================================================================
# HdmiTransition - applies defect + switches resolution (mirrors C# exactly)
# ==============================================================================
TMDS_OFF_TIME_MS = 200

def hdmi_transition(ser, defect_idx, src_idx, cfg):
    jitter_amp  = cfg.jitter_amplitude_mbit
    trans_jit   = cfg.transitional_jitter
    jit_time_ms = cfg.jitter_time_ms
    trans_ip    = cfg.transitional_interpair
    ip_time_ms  = cfg.interpair_delay_ms

    ser.send(",")                     # Disable HDMI
    set_hdmi_source(ser, src_idx)
    #ser.send("!")                     # Reset source

    if defect_idx == 0:   # CLEAN
        time.sleep(TMDS_OFF_TIME_MS / 1000.0)
        ser.send(".")

    elif defect_idx == 1:  # CLOCK_JITTER
        set_clock_jitter(ser, src_idx, jitter_amp)
        if trans_jit:
            time.sleep(jit_time_ms / 1000.0)
            ser.send("W300")          # Remove jitter
        time.sleep(TMDS_OFF_TIME_MS / 1000.0)
        ser.send(".")

    elif defect_idx == 2:  # SYMBOL_SKEW_DELAY
        inter_pair_delay_random(ser, 2, trans_ip, ip_time_ms)
        time.sleep(TMDS_OFF_TIME_MS / 1000.0)
        ser.send(".")

    elif defect_idx == 3:  # SEQUENTIAL_DATA
        sequential_data_random(ser)

    elif defect_idx == 4:  # TMDS_OUTPUT_ENABLE
        time.sleep(TMDS_OFF_TIME_MS / 1000.0)
        glitchy_output_enable_random(ser)

    elif defect_idx == 5:  # TMDS_SCRAMBLING
        time.sleep(TMDS_OFF_TIME_MS / 1000.0)
        scrambling_random(ser)

    elif defect_idx == 6:  # RANDOM_5V_DISCONNECT
        time.sleep(TMDS_OFF_TIME_MS / 1000.0)
        disconnect_5v_random(ser)

    elif defect_idx == 7:  # TMDS_SWAP
        time.sleep(TMDS_OFF_TIME_MS / 1000.0)
        swap_tmds_pair_random(ser)

    elif defect_idx == 8:  # ALL_DRESS
        time.sleep(TMDS_OFF_TIME_MS / 1000.0)
        set_clock_jitter(ser, src_idx, jitter_amp)
        inter_pair_delay_random(ser, 2, trans_ip, ip_time_ms)
        sequential_data_random(ser)
        glitchy_output_enable_random(ser)
        scrambling_random(ser)
        swap_tmds_pair_random(ser)
        swap_tmds_pair_random(ser)

    elif defect_idx == 9:  # BLACK
        time.sleep(TMDS_OFF_TIME_MS / 1000.0)
        # HDMI stays disabled (black screen) - do not call enable


# ==============================================================================
# External tool helpers
# ==============================================================================
def cam_snapshot(cam_index, retries=3):
    """Run CommandCam.exe, return (success, error_msg)."""
    for attempt in range(retries):
        try:
            result = subprocess.run(
                [COMMAND_CAM_EXE, "/devnum", str(cam_index)],
                capture_output=True,
                timeout=30,
                cwd=os.path.dirname(COMMAND_CAM_EXE),
            )
            if result.returncode == 0:
                return True, ""
            err = (result.stdout + result.stderr).decode(errors="replace").strip()
            # kill stale process and retry
            subprocess.run(["taskkill", "/IM", "CommandCam.exe", "/F"],
                           capture_output=True)
            time.sleep(5)
        except FileNotFoundError:
            return False, f"CommandCam.exe not found at {COMMAND_CAM_EXE}"
        except subprocess.TimeoutExpired:
            subprocess.run(["taskkill", "/IM", "CommandCam.exe", "/F"],
                           capture_output=True)
            err = "CommandCam timeout"
    return False, err


def image_processor():
    """Run ImageProcess.exe, return parsed (colorbar%, black%, unknown%) or None.

    Actual output format:
        ImageProcess
        ColorBar  97  0
        Black      2  0
        Unknown    1  0
        process is done!
    """
    try:
        result = subprocess.run(
            [IMAGE_PROCESS_EXE],
            capture_output=True,
            timeout=30,
            cwd=r"C:\HostRhesus",   # image.bmp lives here (written by CommandCam)
        )
        output = result.stdout.decode(errors="replace").strip()
        tokens = output.split()
        # tokens: [ImageProcess, ColorBar, c0, c1, Black, b0, b1, Unknown, u0, u1, process, is, done!]
        # indices:      0           1       2   3    4     5   6     7      8   9
        colorbar = int(tokens[2])
        black    = int(tokens[5])
        unknown  = int(tokens[8])
        return colorbar, black, unknown
    except Exception:
        return None


def ping_dut(ip):
    """Ping the DUT once, return True if reply received."""
    try:
        result = subprocess.run(
            ["ping.exe", "-n", "1", ip],
            capture_output=True,
            timeout=10,
        )
        output = result.stdout.decode(errors="replace")
        lines = [l.strip() for l in output.splitlines() if l.strip()]
        if len(lines) > 1 and lines[1].startswith("Reply"):
            return True
    except Exception:
        pass
    return False


# ==============================================================================
# Next value selector (INCREMENTAL / RANDOM / STATIC with mask)
# ==============================================================================
def get_next_value(algo, current, count, mask):
    if (((1 << count) - 1) & mask) == 0:
        return current  # no bits set, stay put
    nxt = current
    for _ in range(count * 2):  # safety loop limit
        if algo == "INCREMENTAL":
            nxt = (nxt + 1) % count
        elif algo == "RANDOM":
            nxt = random.randint(0, count - 1)
        elif algo == "STATIC":
            return current
        if (mask >> nxt) & 1:
            return nxt
    return current


# ==============================================================================
# Config object
# ==============================================================================
class Config:
    pass


# ==============================================================================
# Build ordered list of enabled indices from a bitmask
# ==============================================================================
def enabled_indices(mask, count):
    """Return list of indices where the corresponding bit is set in mask."""
    return [i for i in range(count) if (mask >> i) & 1]


# ==============================================================================
# Run one iteration (single resolution + single defect combination)
# ==============================================================================
def run_iteration(ser, src_idx, defect_idx, iteration_num, cfg, logger,
                  pass_cnt, fail_cnt, nocheck_cnt, monitor_err):
    """Execute one resolution+defect combination. Returns updated counters."""

    ts = datetime.datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
    line = f"{ts} {iteration_num:5d} D{defect_idx} "
    line += source_info_str(src_idx)
    logger.log_inline(line)

    init_registers(ser)
    hdmi_transition(ser, defect_idx, src_idx, cfg)

    # Wait before snapshot in 100ms steps
    for _ in range(cfg.loop_delay_ms // 100):
        time.sleep(0.1)

    check_image = ((cfg.defect_image_mask >> defect_idx) & 1) != 0

    colorbar_pct = 0
    black_pct    = 0
    unknown_pct  = 0
    snap_fatal   = False

    if check_image:
        snap_ok, snap_err = cam_snapshot(cfg.cam_index)
        if snap_ok:
            logger.log_inline(" shoot OK  ")
        else:
            logger.log_inline(f" shoot FAIL: {snap_err}")
            if "not found" in snap_err:
                snap_fatal = True

        if snap_ok:
            recog = image_processor()
            if recog:
                colorbar_pct, black_pct, unknown_pct = recog
                logger.log_inline(
                    f" Recog c{colorbar_pct:3d} b{black_pct:3d} u{unknown_pct:3d}"
                )
            else:
                logger.log_inline(" Recog FAIL")
    else:
        logger.log_inline(" shoot --   Recog c -- b -- u --")

    if cfg.ping_enable and cfg.ping_ip:
        ping_ok = ping_dut(cfg.ping_ip)
        logger.log_inline(" Ping OK  " if ping_ok else " Ping FAIL")

    # Evaluate pass/fail
    is_black_source = ((cfg.source_black_mask >> src_idx) & 1) != 0
    eval_value = black_pct if (defect_idx == 9 or is_black_source) else colorbar_pct

    ping_ok_val = True
    if cfg.ping_enable and cfg.ping_ip:
        ping_ok_val = ping_dut(cfg.ping_ip)

    if (cfg.ping_enable and not ping_ok_val) or \
       (check_image and eval_value < cfg.image_threshold):
        fail_cnt    += 1
        monitor_err += 1
        result_str = f" FAIL (Pass {pass_cnt:5d}, NoCheck {nocheck_cnt:5d}, Fail {fail_cnt:5d}, total {iteration_num:5d})"
    elif not check_image:
        nocheck_cnt += 1
        result_str = f" ---- (Pass {pass_cnt:5d}, NoCheck {nocheck_cnt:5d}, Fail {fail_cnt:5d}, total {iteration_num:5d})"
    else:
        pass_cnt += 1
        result_str = f" PASS (Pass {pass_cnt:5d}, NoCheck {nocheck_cnt:5d}, Fail {fail_cnt:5d}, total {iteration_num:5d})"

    logger.log(result_str)

    # After logging result — check API and re-lock if Avio2 lost the stream
    check_and_relock_if_needed(ser, src_idx, colorbar_pct, black_pct, defect_idx, logger)

    return pass_cnt, fail_cnt, nocheck_cnt, monitor_err, snap_fatal

def get_avio2_token():
    """Get Bearer token from Avio2 TX API."""
    import requests
    import urllib3
    urllib3.disable_warnings()
    try:
        r = requests.post(
            AVIO2_AUTH_URL,
            auth=(AVIO2_USERNAME, AVIO2_PASSWORD),
            verify=False,
            timeout=10
        )
        if r.status_code == 200:
            return r.json().get("accessToken")
    except Exception as e:
        print(f"[API] Auth failed: {e}")
    return None


def check_and_relock_if_needed(ser, src_idx, colorbar_pct, black_pct, defect_idx, logger):
    urllib3.disable_warnings()

    is_black_defect = defect_idx == 9
    if black_pct < 90 or is_black_defect:
        return False

    logger.log(f"\n[API] Colorbar lost (b={black_pct}), checking Avio2 stream state...")

    token = get_avio2_token()
    if not token:
        logger.log("[API] Could not get auth token, skipping API check")
        return False

    headers = {"Authorization": f"Bearer {token}"}
    try:
        r = requests.get(AVIO2_STREAM_URL, headers=headers, verify=False, timeout=10)
        if r.status_code == 200:
            response_json = r.json()
            # Log full response so we can see actual structure
            logger.log(f"[API] Full stream response: {response_json}")
            state = response_json.get("state")
            logger.log(f"[API] Stream state value: {state!r} (expected {STREAM_STATE_ACTIVE} for Active)")

            if state is None:
                logger.log("[API] Warning: 'state' field not found in response — check AVIO2_STREAM_URL endpoint")
                return False

            if state != STREAM_STATE_ACTIVE:
                logger.log("[API] Stream is Inactive — performing re-lock to 1080P60...")
                # Full clean re-lock sequence
                disable_hdmi(ser)
                time.sleep(2)
                init_registers(ser)
                set_hdmi_source(ser, 11)    # HDMI_1920x1080P60
                ser.send("!")
                time.sleep(3)
                enable_hdmi(ser)
                logger.log(f"[API] Waiting {RELOCK_WAIT_S}s for Avio2 to re-lock...")
                time.sleep(RELOCK_WAIT_S)

                # Verify recovery by checking stream state again
                r2 = requests.get(AVIO2_STREAM_URL, headers=headers, verify=False, timeout=10)
                if r2.status_code == 200:
                    state2 = r2.json().get("state")
                    if state2 == STREAM_STATE_ACTIVE:
                        logger.log("[API] Stream confirmed Active after re-lock. Resuming test.")
                    else:
                        logger.log(f"[API] Warning: Stream state still {state2!r} after re-lock.")
                return True
            else:
                logger.log("[API] Stream is Active per API — colorbar loss may be transient, continuing.")
        else:
            logger.log(f"[API] Stream status check failed: HTTP {r.status_code}")
    except Exception as e:
        logger.log(f"[API] Stream check error: {e}")

    return False


# ==============================================================================
# Main test loop  — nested: outer = resolution, inner = defect
# ==============================================================================
def run_test(cfg, logger):
    # Build ordered lists of enabled sources and defects from masks
    src_list    = enabled_indices(cfg.source_mask, SOURCE_COUNT)
    defect_list = enabled_indices(cfg.defect_mask, DEFECT_COUNT)

    if not src_list:
        logger.log("FATAL: No resolutions enabled (source_mask=0). Stopping.")
        return 1
    if not defect_list:
        logger.log("FATAL: No defects enabled (defect_mask=0). Stopping.")
        return 1

    # Count how many iterations one full matrix pass takes
    matrix_size = len(src_list) * len(defect_list)

    logger.log(f"\nTest started at: {datetime.datetime.now()}")
    logger.log(f"Logging data to file: {cfg.log_file}")
    logger.log("HDMI Monitor test with Nexys Source")
    logger.log(f"Script: hdmi_gen_test.py")
    logger.log(f"Test mode: NESTED (all defects per resolution)")
    logger.log(f"Enabled resolutions ({len(src_list)}): "
               f"{', '.join(SOURCE_NAMES[i] for i in src_list)}")
    logger.log(f"Enabled defects    ({len(defect_list)}): "
               f"{', '.join(DEFECT_NAMES[i] for i in defect_list)}")
    logger.log(f"Matrix size: {len(src_list)} resolutions x {len(defect_list)} defects "
               f"= {matrix_size} iterations per pass")
    loops_label = f"{cfg.max_loops} passes" if cfg.max_loops > 0 else "unlimited passes"
    logger.log(f"Passes requested: {loops_label}  "
               f"({'%d' % (matrix_size * cfg.max_loops)} total iterations)" if cfg.max_loops > 0
               else f"Passes requested: unlimited")
    logger.log(f"Test parameters : source_mask=0x{cfg.source_mask:04X}, "
               f"defect_mask=0x{cfg.defect_mask:04X}, "
               f"image_mask=0x{cfg.defect_image_mask:04X}")
    logger.log(f"Test parameters : StopAfterErrorCount={cfg.stop_after_error}, "
               f"Defect_Jitter_mTbit={cfg.jitter_amplitude_mbit}, "
               f"loop_delay_ms={cfg.loop_delay_ms}")

    # Open serial port
    try:
        ser = HdmiGenSerial(cfg.com_port)
    except Exception as e:
        logger.log(f"\nFATAL: Cannot open {cfg.com_port}: {e}")
        sys.exit(1)

    logger.log("Init Source HDMI.")
    disable_hdmi(ser)
    init_registers(ser)
    set_hdmi_source(ser, src_list[0])
    time.sleep(1)
    enable_hdmi(ser)
    time.sleep(5)

    # Force 1080P60 on startup to ensure Avio2 locks on before main loop begins
    logger.log("Forcing 1920x1080P60 startup resolution to initialize Avio2...")
    init_registers(ser)
    set_hdmi_source(ser, 11)    # HDMI_1920x1080P60
    ser.send("!")               # Reset
    time.sleep(3)
    enable_hdmi(ser)
    logger.log("Waiting 15s for Avio2 to lock onto 1080P60 signal...")
    time.sleep(15)
    logger.log("Avio2 should now be active. Starting main test loop.")

    iteration_num = 0
    pass_cnt      = 0
    fail_cnt      = 0
    nocheck_cnt   = 0
    monitor_err   = 0
    stop          = False
    pass_num      = 0  # how many full matrix passes completed

    try:
        while True:
            pass_num += 1
            logger.log(f"\n--- Pass {pass_num}"
                       + (f" of {cfg.max_loops}" if cfg.max_loops > 0 else "") + " ---")

            for src_idx in src_list:
                if stop:
                    break
                logger.log(f"\n  >> Resolution: {SOURCE_NAMES[src_idx]}")

                for defect_idx in defect_list:
                    if stop:
                        break

                    iteration_num += 1
                    pass_cnt, fail_cnt, nocheck_cnt, monitor_err, snap_fatal = run_iteration(
                        ser, src_idx, defect_idx, iteration_num, cfg, logger,
                        pass_cnt, fail_cnt, nocheck_cnt, monitor_err
                    )

                    if snap_fatal:
                        logger.log("\nFATAL: CommandCam.exe not found. Stopping.")
                        stop = True
                        break

                    if cfg.stop_after_error > 0 and fail_cnt >= cfg.stop_after_error:
                        logger.log(f"\nStopping after {fail_cnt} error(s) as configured.")
                        stop = True
                        break

            if stop:
                break

            # Check if we've completed the requested number of passes
            if cfg.max_loops > 0 and pass_num >= cfg.max_loops:
                logger.log(f"\nAll {cfg.max_loops} pass(es) complete. Stopping.")
                break

    except KeyboardInterrupt:
        logger.log("\nTest interrupted by user (Ctrl+C).")

    finally:
        ser.close()

    # Final report
    error_rate = (100 * monitor_err / iteration_num) if iteration_num > 0 else 0.0
    logger.log(f"\nPasses completed: {pass_num}")
    logger.log(f"Iterations,                                    TOTAL COUNT = {iteration_num}")
    logger.log(f"Monitor Fail                                   ERROR COUNT = {monitor_err}    ERROR RATE = {error_rate:.2f} %")
    logger.log(f"Pass: {pass_cnt}   NoCheck: {nocheck_cnt}   Fail: {fail_cnt}   Total: {iteration_num}")
    logger.log(f"Test complete at: {datetime.datetime.now()}")

    return 1 if fail_cnt > 0 else 0


# ==============================================================================
# CLI argument parser
# ==============================================================================
def parse_args():
    parser = argparse.ArgumentParser(
        description="HDMI Generator Test - Nexys Video FPGA board controller"
    )
    parser.add_argument("--com-port",         default=DEFAULT_COM_PORT,
                        help=f"Serial COM port (default: {DEFAULT_COM_PORT})")
    parser.add_argument("--cam-index",        type=int, default=DEFAULT_CAM_INDEX,
                        help=f"CommandCam webcam device index (default: {DEFAULT_CAM_INDEX})")
    parser.add_argument("--log-file",         default=DEFAULT_LOG_FILE,
                        help=f"Output log file path (default: {DEFAULT_LOG_FILE})")
    parser.add_argument("--max-loops",        type=int, default=0,
                        help="Max iterations (0 = run forever until Ctrl+C)")
    parser.add_argument("--stop-after-error", type=int, default=DEFAULT_STOP_AFTER_ERROR,
                        help="Stop after N errors (0 = never stop)")
    parser.add_argument("--loop-delay-ms",    type=int, default=DEFAULT_LOOP_DELAY_MS,
                        help=f"ms to wait before webcam snapshot (default: {DEFAULT_LOOP_DELAY_MS})")
    parser.add_argument("--source-algo",      default=DEFAULT_SOURCE_ALGO,
                        choices=["INCREMENTAL", "RANDOM", "STATIC"],
                        help="Source change algorithm")
    parser.add_argument("--source-start",     type=int, default=DEFAULT_SOURCE_START,
                        help=f"Source start index 0-{SOURCE_COUNT-1} (default: {DEFAULT_SOURCE_START})")
    parser.add_argument("--source-mask",      type=lambda x: int(x, 0),
                        default=DEFAULT_SOURCE_MASK,
                        help=f"Source enable bitmask (default: 0x{DEFAULT_SOURCE_MASK:04X})")
    parser.add_argument("--source-black-mask",type=lambda x: int(x, 0),
                        default=DEFAULT_SOURCE_BLACK_MASK,
                        help="Source black screen mask")
    parser.add_argument("--defect-algo",      default=DEFAULT_DEFECT_ALGO,
                        choices=["INCREMENTAL", "RANDOM", "STATIC"],
                        help="Defect change algorithm")
    parser.add_argument("--defect-start",     type=int, default=DEFAULT_DEFECT_START,
                        help=f"Defect start index 0-{DEFECT_COUNT-1} (default: {DEFAULT_DEFECT_START})")
    parser.add_argument("--defect-mask",      type=lambda x: int(x, 0),
                        default=DEFAULT_DEFECT_MASK,
                        help=f"Defect enable bitmask (default: 0x{DEFAULT_DEFECT_MASK:04X})")
    parser.add_argument("--defect-image-mask",type=lambda x: int(x, 0),
                        default=DEFAULT_DEFECT_IMAGE_MASK,
                        help="Image check mask per defect type")
    parser.add_argument("--jitter-amplitude", type=int, default=DEFAULT_JITTER_AMPLITUDE_MBIT,
                        help=f"Jitter amplitude in mTbit (default: {DEFAULT_JITTER_AMPLITUDE_MBIT})")
    parser.add_argument("--image-threshold",  type=int, default=DEFAULT_IMAGE_THRESHOLD,
                        help=f"Min colorbar %% to call PASS (default: {DEFAULT_IMAGE_THRESHOLD})")
    parser.add_argument("--ping-ip",          default=DEFAULT_PING_IP,
                        help="DUT IP to ping each iteration (empty = no ping)")
    parser.add_argument("--no-transitional-jitter", action="store_true",
                        help="Keep jitter permanently (do not remove after jitter_time_ms)")
    parser.add_argument("--jitter-time-ms",   type=int, default=DEFAULT_JITTER_TIME_MS,
                        help=f"How long jitter is applied ms (default: {DEFAULT_JITTER_TIME_MS})")
    parser.add_argument("--no-transitional-interpair", action="store_true",
                        help="Keep inter-pair delay permanently")
    parser.add_argument("--interpair-delay-ms",type=int, default=DEFAULT_INTERPAIR_DELAY_MS,
                        help=f"How long inter-pair delay is applied ms (default: {DEFAULT_INTERPAIR_DELAY_MS})")
    return parser.parse_args()


# ==============================================================================
# Entry point
# ==============================================================================
def main():
    args = parse_args()

    cfg = Config()
    cfg.com_port               = args.com_port
    cfg.cam_index              = args.cam_index
    cfg.log_file               = args.log_file
    cfg.max_loops              = args.max_loops
    cfg.stop_after_error       = args.stop_after_error
    cfg.loop_delay_ms          = args.loop_delay_ms
    cfg.source_algo            = args.source_algo
    cfg.source_start           = args.source_start
    cfg.source_mask            = args.source_mask
    cfg.source_black_mask      = args.source_black_mask
    cfg.defect_algo            = args.defect_algo
    cfg.defect_start           = args.defect_start
    cfg.defect_mask            = args.defect_mask
    cfg.defect_image_mask      = args.defect_image_mask
    cfg.jitter_amplitude_mbit  = args.jitter_amplitude
    cfg.image_threshold        = args.image_threshold
    cfg.ping_ip                = args.ping_ip
    cfg.ping_enable            = bool(args.ping_ip)
    cfg.transitional_jitter    = not args.no_transitional_jitter
    cfg.jitter_time_ms         = args.jitter_time_ms
    cfg.transitional_interpair = not args.no_transitional_interpair
    cfg.interpair_delay_ms     = args.interpair_delay_ms

    logger = Logger(cfg.log_file)
    exit_code = run_test(cfg, logger)
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
