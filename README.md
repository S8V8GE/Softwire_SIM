# 🚪 Softwire Simulator (Softwire_SIM)

> Proof of concept for simulating realistic access control behaviour in Softwire

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Status](https://img.shields.io/badge/status-beta-orange)
![Purpose](https://img.shields.io/badge/purpose-simulation-lightgrey)

---

## 🧠 Overview

Softwire_SIM is a PowerShell-based tool that simulates real-world door activity within a Softwire environment.

It allows you to both generate controlled and simulate random access control events without requiring physical interaction with hardware.

The `Alpha` release will be wrapped in an interactive GUI and replace `SimNG` in EMEA STC-001 and STC-002 trainings.

---

## ⚙️ Features

- 🔁 Automated door activity simulation 
- 🎯 Configurable event types (Normal / Forced / Held)  
- ⏱️ Adjustable timing between events  
- 🚪 Intelligent door selection & filtering  
- 🧪 Safe testing without impacting real systems  

---

## 🧪 Use Cases

- Training & demonstrations  
- Testing alarm behaviour  
- Validating monitoring workflows  
- Reproducing edge cases (forced / held doors)  

---

## 📦 Requirements

- PowerShell 5.1+ (PowerShell 7 recommended)
- Softwire environment
- `SoftwirePSM` module - Genetec employees only

---

## 🚀 Getting Started

```powershell
git clone https://github.com/S8V8GE/Softwire_SIM.git
cd Softwire_SIM

Import-Module SoftwirePSM
.\Softwire_BETA_v1.ps1

---

## ✅ What Works

### 🟢 Basic Configuration

- **Auto-enrol via door**
  - Present card prompt → simulate read → credential is enrolled successfully

- **Door states**
  - Door forced  
  - Door held open too long  

- **Access Granted**
  - Without door sensor → relay follows grant time (standard & extended ✔)
  - With door sensor → relay follows relock rules (on close / timed ✔)

- **Reader behaviour**
  - Read In / Read Out ✔  
  - REX (all modes) ✔  
  - Card + PIN ✔  
  - PIN only ✔  
  *(Requires "Card or PIN" enabled in Unit Wide Parameters)*

- **Access Denied**
  - Works as expected ✔

- **Schedules**
  - Unlock schedules ✔  
  - Exceptions ✔  

- **Door settings**
  - Ignore "Door open too long" ✔  
  - Ignore "Access granted/denied" ✔  

- **Security Desk Door Widget**
  - Reader shunting ✔  
  - Maintenance mode ✔  
  - Override unlock schedules ✔  
  - Input shunting ✔  

- **Area Presence**
  - Movement tracking works correctly ✔  
  - Reports & Access Troubleshooter ✔  

---

### 🔵 Advanced Configuration

- **Two-person rule** ✔  

- **Anti-passback**
  - Soft / Soft + Strict ✔  
  - Hard / Hard + Strict ✔  
  - Presence timeout ✔  
  - Bypass antipassback ✔  
  - ⚠️ Only works properly on perimeter doors  
  - ✅ *Forgive antipassback violation now works*

- **Max Occupancy**
  - Soft ✔  
  - Hard ✔  
  - Bypass antipassback ✔  

- **Door Interlock**
  - Interlock ✔  
  - Override ✔  
  - Lockdown ✔  

- **First Person In Rule**
  - Enforced on schedules ✔  
  - Enforced on access rules ✔  

- **Visitor Escort**
  - Standard escort ✔  
  - Single passage ✔  
  *(Remove OUT reader as per technote)*

- **Duress PIN** ✔  

- **Double Badge**
  - Single cardholder ✔  
  - Cardholder groups ✔  

- **Threat Levels**
  - Clearance levels ✔  
  - Lockdown scenarios ✔  
  - Fire scenarios ✔  

---

## ❌ Limitations

- **Entry Sensor**
  - "No entry detected" event could not be triggered

- **Lock Sensor**
  - Not supported (not exposed via Softwire API)

- **Buzzer**
  - Not supported (not exposed via Softwire API)







