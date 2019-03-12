import requests
import subprocess
import time
import sched
import xml.etree.ElementTree as xmlET
import configparser
from pyautogui import press

parser = configparser.ConfigParser()
#parser.read('C:\\Users\\user\\Desktop\\config.INI')
parser.read('C:\\Users\\Morgan.Rehnberg\\Desktop\\config.INI')
config = parser['Config']

name = config['name']
ip = 'localhost'
prefix = "http://"
postfix = ":5050/layerApi.aspx?cmd="
httpSession = requests.Session()

idle = False
last_idle_check_state = {'lat': 0, 'lon': 0, 'zoom': 0}
idle_t = 30 # Interval in seconds to check for idle

old_spin_state = {}
spin_t = .5 # Interval in seconds to check for spin. Should be fast.

min_zoom = config.getfloat('min_zoom')
max_zoom = config.getfloat('max_zoom')

movement_block = False # This is set when motion is begun to keep the spin checker from freaking out
startup_block = False # This is set if we restart WWT to give it time to start

# Create an event scheduler
s = sched.scheduler()

def reset_movement_block():

    global movement_block
    movement_block = False

def reset_startup_block():

    # This rests the block on trying a new startup of WWT. It also sends a keypress of F11 to put everything fullscreen.
    
    global startup_block
    startup_block = False
    press('f11')

def check_for_idle():

    global idle
    global last_idle_check_state
    
    print('Checking for idle... ', end="")
    result = get_state()
    if result is None:
        s.enter(idle_t, 3, check_for_idle)
        return
    
    try:
        root = xmlET.fromstring(result.text)
        props = root[1].attrib
        lat = props['lat']
        lon = props['lng']
        zoom = props['zoom']
    except:
        setup()
        s.enter(idle_t, 3, check_for_idle)
        return
    
    old = last_idle_check_state
    
    if (lat == old['lat']) and (lon == old['lon']) and (zoom == old['zoom']):
        print('Idle')
        if not idle: # Don't issue commands if we already know we're idle
            idle = True
            # Only reset if the planet is too large or small. Ignore rotation changes.
            if not ((float(zoom) >= min_zoom) and (float(zoom) <= max_zoom)):
                print(float(zoom) > 1e-6)
                print(float(zoom) < .0004)
                print(zoom)
                setup()
            else:
                print('Zoom range acceptable; no setup')
    else:
        idle = False
        print('Not idle')   
        
    last_idle_check_state['lat'] = lat
    last_idle_check_state['lon'] = lon
    last_idle_check_state['zoom'] = zoom
    # Schedule the event to run again in idle_t seconds
    s.enter(idle_t, 3, check_for_idle)

def rapid_check():

    # Function that checks for various situations at a high cadence. Checks:
    # - That WWT is running
    # - That the correct object is on the screen
    # - That the object is not stuck in a rapid spin.
    
    global movement_block
    global startup_block
    
    # Check if WWT is running. If not, restart it.
    active = check_WWT_health()
    if not active:
        if not startup_block: # We haven't already tried to startup recently
            startup_block = True
            s.enter(30,4, reset_startup_block)
            subprocess.Popen("C:\\Program Files (x86)\\Microsoft Research\\Microsoft WorldWide Telescope\\WWTExplorer.exe")
            print('WWT down! Attempting restart...')

    result = get_state()

    if result is None:
        s.enter(spin_t, 2, rapid_check)
        return
     
    try:
        root = xmlET.fromstring(result.text)
        props = root[1].attrib
    except: # Bad XML usually means we're in the wrong mode
        setup()
        s.enter(spin_t, 2, rapid_check)
        return
    
    # Check if they've navigated away
    check_for_wrong_object(props)
    
    # Check for rapid spin
    try:
        lat = float(props['lat'])
        lon = float(props['lng'])
    except:
        setup()
        s.enter(spin_t, 2, rapid_check)
        return
        
    # We don't check for spin while the camera is being moved by setup(), but we do keep track of the changing state.
    if not movement_block:
        if 'lat' in old_spin_state: 
            # Check for > 100 deg/sec rotation, taking care to deal with the 360 -> 0 boundary
            dLat = (lat - old_spin_state['lat'] + 180) % 360 -180
            dLon = (lon - old_spin_state['lon'] + 180) % 360 -180
            
            if (abs(dLat) > 100) or (abs(dLon) > 100):
                print('Fixing spin')
                setup()
                movement_block = True
                s.enter(15, 1, reset_movement_block)
        else:
            old_spin_state['lat'] = lat
            old_spin_state['lon'] = lon
    else:
        old_spin_state['lat'] = lat
        old_spin_state['lon'] = lon
    
    s.enter(1,2, rapid_check) # Re-queue the next call

def check_WWT_health():

    # Function to look for whether the WWT process is active. Returns True if WWT is running.

    name = 'WWTExplorer.exe'
    cmd = 'tasklist /FI "IMAGENAME eq %s"' % name
    status = subprocess.Popen(cmd, stdout=subprocess.PIPE).stdout.read()
    return name in str(status)
    
def check_for_wrong_object(state):

    # Function to check whether the user has switched to another object
    
    frame = state['ReferenceFrame']
    name = config.get('name')
    
    if name != 'Saturn':
        if name != frame:
            print('Frame changed! Resetting...')
            setup()
    else:
        if frame != 'Sun':
            setup()
            print('Frame changed! Resetting...')
    
def setup():
    print ('setup')
    
    global movement_block
    
    movement_block = True
    s.enter(15, 1, reset_movement_block)
    
    flyto = config.get('flyto_command')
    zoom = config.get('zoom_command')
    try:
        req = httpSession.post(prefix+ip+postfix+flyto)
        req2 = httpSession.post(prefix+ip+postfix+zoom)
    except:
        return
    
def get_state():
    command = "state"
    try:
        req = httpSession.post(prefix+ip+postfix+command, timeout=0.05)
    except:
        req = None

    return(req)

def launch_WWT():
    
    subprocess.Popen("C:\Program Files (x86)\Microsoft Research\Microsoft WorldWide Telescope\WWTExplorer.exe")
    
print('Setting up the screen...')
setup()
# Check whether the instance is idle every idle_t seconds
s.enter(idle_t, 3, check_for_idle)
# Check whether the planet is spinning every spin_t seconds
s.enter(spin_t, 2, rapid_check)
s.run()