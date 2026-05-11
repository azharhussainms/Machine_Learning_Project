#!/usr/bin/env python3
"""Convert the proposal DOCX files to PDF using LibreOffice's UNO bridge,
updating the Table of Contents and all fields (PAGE / NUMPAGES) before export."""
import os, sys, time, subprocess, signal

DOCS = ["Proposal_1.docx", "Proposal_2.docx", "Proposal_3.docx", "Master_Proposal_Document.docx"]
PORT = 2199

def file_url(path):
    return "file://" + os.path.abspath(path)

def main():
    # launch headless soffice listener
    proc = subprocess.Popen(
        ["soffice", "--headless", "--norestore", "--nologo", "--nofirststartwizard",
         f"--accept=socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        import uno
        from com.sun.star.beans import PropertyValue

        def prop(name, value):
            p = PropertyValue(); p.Name = name; p.Value = value; return p

        # connect (retry until ready)
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext)
        ctx = None
        for _ in range(40):
            try:
                ctx = resolver.resolve(
                    f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.75)
        if ctx is None:
            print("ERROR: could not connect to LibreOffice"); sys.exit(1)
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        for d in DOCS:
            if not os.path.exists(d):
                print("skip (missing):", d); continue
            doc = desktop.loadComponentFromURL(file_url(d), "_blank", 0,
                                               (prop("Hidden", True),))
            try:
                # dispatcher for .uno: commands (needs the document's frame)
                dispatcher = smgr.createInstanceWithContext(
                    "com.sun.star.frame.DispatchHelper", ctx)
                frame = doc.getCurrentController().getFrame()
                def call(cmd):
                    try:
                        dispatcher.executeDispatch(frame, cmd, "", 0, ())
                    except Exception as e:
                        print("  dispatch warn", cmd, e)
                # update everything: links, fields, charts, indexes
                call(".uno:UpdateAllLinks")
                call(".uno:UpdateCharts")
                call(".uno:UpdateAllIndexes")
                call(".uno:UpdateInputFields")
                call(".uno:UpdateFields")
                call(".uno:UpdateAll")
                try:
                    doc.refresh()
                    doc.getTextFields().refresh()
                except Exception:
                    pass
                try:
                    idxs = doc.getDocumentIndexes()
                    n = idxs.getCount()
                    for i in range(n):
                        idxs.getByIndex(i).update()
                    print(f"  indexes found/updated: {n}")
                except Exception as e:
                    print("  index update warning:", e)
                # second pass so page numbers settle after the TOC grew
                call(".uno:UpdateAllIndexes")
                call(".uno:UpdateAll")
                try:
                    doc.refresh(); doc.getTextFields().refresh()
                except Exception:
                    pass
                pdf_path = os.path.splitext(d)[0] + ".pdf"
                doc.storeToURL(file_url(pdf_path), (prop("FilterName", "writer_pdf_Export"),))
                print("exported:", pdf_path)
            finally:
                doc.close(False)
    finally:
        try:
            proc.terminate(); proc.wait(timeout=10)
        except Exception:
            proc.kill()

if __name__ == "__main__":
    main()
