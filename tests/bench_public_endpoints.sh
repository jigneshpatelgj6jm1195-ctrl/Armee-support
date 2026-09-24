#!/bin/bash
# Read-only latency benchmark for the public Apps Script endpoints (no writes).
# usage: tests/bench_public_endpoints.sh <web-app /exec URL> [runs]
# Run it against the current deployment and a test deployment with the same
# runs count to compare before/after under comparable conditions.
U=$1; N=${2:-8}; D=24180700110
for a in "health" "get_master_version" "get_master" "get_school_updates" "get_school_complaints&dise=$D" "get_department_complaint&dise=$D" "check_duplicate&serial=ZZBENCH0001"; do
  ts=(); for i in $(seq $N); do ts+=($(curl -sSL -o /dev/null -w "%{time_total}" "$U?action=$a&t=$RANDOM$i")); done
  printf "%s\n" "${ts[@]}" | sort -n | awk -v a="$a" '{v[NR]=$1} END{printf "%-45s median=%.2fs p90=%.2fs min=%.2fs max=%.2fs n=%d\n", a, v[int((NR+1)/2)], v[int(NR*0.9+0.5)], v[1], v[NR], NR}'
done
