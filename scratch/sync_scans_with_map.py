import os
import sys
# Add parent directory to path to ensure app/models can be imported
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dotenv import load_dotenv
load_dotenv()

from app import app, db
from models import Scan, CustomerMap, Cylinder

def run_sync():
    with app.app_context():
        print("Starting scan-to-mapping database synchronization...")
        cmaps = CustomerMap.query.all()
        updated_scans = 0
        updated_cyls = 0
        
        for cmap in cmaps:
            # Find matching scans
            scans = Scan.query.filter_by(
                scan_date=cmap.scan_date,
                scan_time=cmap.scan_time,
                driver=cmap.driver,
                action=cmap.action
            ).all()
            
            for s in scans:
                if s.customer != cmap.customer:
                    print(f"Syncing scan {s.id} ({s.cylinder_uid}): {s.customer} -> {cmap.customer}")
                    s.customer = cmap.customer
                    updated_scans += 1
            
            # If this is a Delivery, ensure cylinder registry is updated
            if cmap.action == 'Delivery' and cmap.customer:
                uids = [s.cylinder_uid for s in scans]
                for uid in uids:
                    cyl = Cylinder.query.filter(Cylinder.uid.ilike(uid)).first()
                    if cyl and cyl.location != cmap.customer:
                        print(f"Syncing cylinder location {cyl.uid}: {cyl.location} -> {cmap.customer}")
                        cyl.location = cmap.customer
                        updated_cyls += 1
                        
        db.session.commit()
        print(f"Sync complete! Updated {updated_scans} scan logs and {updated_cyls} cylinder registry locations.")

if __name__ == "__main__":
    run_sync()
