# 🚪 Softwire Simulator (Softwire_SIM)

> Proof of concept for simulating realistic access control behaviour in Softwire

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Status](https://img.shields.io/badge/status-beta-orange)
![Purpose](https://img.shields.io/badge/purpose-simulation-lightgrey)

---

## 🧠 Overview

Softwire_SIM is a PowerShell-based tool that simulates real-world door activity within a Softwire environment.

It allows you to generate controlled access control events without requiring physical interaction with hardware.

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
- `SoftwirePSM` module

---

## 🚀 Getting Started

```powershell
git clone https://github.com/S8V8GE/Softwire_SIM.git
cd Softwire_SIM

Import-Module SoftwirePSM
.\Softwire_BETA_v1.ps1
