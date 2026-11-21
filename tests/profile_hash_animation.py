"""Measure frame scheduling during real verification at two file counts."""
import json
import statistics
import time
import test_review_fixes as fixtures
from PySide6.QtCore import QTimer, Qt

m, app = fixtures.m, fixtures.app
case = fixtures.ReviewFixTests('test_all_category_extensions_remain_reachable')
case.setUp()
window = case.make_window()
results = []
try:
    case.sample.write_bytes(b'harmless verification fixture' * 2048)
    window._switch_to_hash_mode()
    window.show()
    for count in (200, 2000):
        window.hash_items = [dict(id=str(i), name=f'Fixture {i}.dat', path=str(case.sample),
            size=case.sample.stat().st_size, algorithm='SHA-256', expected_hash='', status='pending',
            status_text='Pending', result='PENDING') for i in range(count)]
        frames = []
        def sample():
            if window.verification_loader.state == 'verifying':
                frames.append((time.monotonic(), window.verification_loader.angle))
        window.verification_loader.timer.timeout.connect(sample)
        window._start_hash_verification()
        deadline = time.monotonic() + 30
        while not window.btn_verify_hashes.isEnabled() and time.monotonic() < deadline:
            app.processEvents()
            time.sleep(.001)
        window.verification_loader.timer.timeout.disconnect(sample)
        window.hash_worker.wait(5000)
        assert window.btn_verify_hashes.isEnabled()
        assert all(item['result'] == 'NO TARGET' for item in window.hash_items)
        intervals = sorted((frames[i][0] - frames[i-1][0]) * 1000 for i in range(1, len(frames)))
        assert len(intervals) > 2
        rotation = sum((frames[i][1] - frames[i-1][1]) % 360 for i in range(1, len(frames)))
        speed = rotation / (frames[-1][0] - frames[0][0])
        results.append(dict(files=count, frames=len(frames), median_frame_ms=round(statistics.median(intervals), 2),
            p95_frame_ms=round(intervals[int((len(intervals)-1)*.95)], 2), max_frame_ms=round(max(intervals), 2),
            measured_degrees_per_second=round(speed, 2)))
    print(json.dumps(results, indent=2))
    (fixtures.ROOT / 'scratch/hash_animation_profile.json').write_text(json.dumps(results, indent=2), encoding='utf-8')
finally:
    window.close()
    case.tearDown()
    case.doCleanups()
