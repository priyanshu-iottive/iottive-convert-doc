import { useState, useCallback, useRef } from "react";
import { Card, CardContent } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { useToast } from "@/hooks/use-toast";
import { Upload, FileText, Download, Loader2, X, CheckCircle2 } from "lucide-react";
import { apiRequest } from "@/lib/queryClient";

type ConversionState = "idle" | "uploading" | "converting" | "done" | "error";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [clientName, setClientName] = useState("");
  const [contactName, setContactName] = useState("");
  const [contactEmail, setContactEmail] = useState("");
  const [projectTitle, setProjectTitle] = useState("");
  const [state, setState] = useState<ConversionState>("idle");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [downloadName, setDownloadName] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [errorMsg, setErrorMsg] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const { toast } = useToast();

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const dropped = e.dataTransfer.files[0];
    if (dropped) {
      const ext = dropped.name.split(".").pop()?.toLowerCase();
      if (ext === "pdf" || ext === "docx") {
        setFile(dropped);
        setState("idle");
        setDownloadUrl(null);
      } else {
        toast({
          title: "Unsupported format",
          description: "Please upload a PDF or DOCX file.",
          variant: "destructive",
        });
      }
    }
  }, [toast]);

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selected = e.target.files?.[0];
    if (selected) {
      setFile(selected);
      setState("idle");
      setDownloadUrl(null);
    }
  };

  const removeFile = () => {
    setFile(null);
    setState("idle");
    setDownloadUrl(null);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const handleConvert = async () => {
    if (!file) return;

    setState("uploading");
    setErrorMsg("");

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("clientName", clientName || "Client Name");
      formData.append("contactName", contactName || "Contact Person");
      formData.append("contactEmail", contactEmail || "email@example.com");
      formData.append("projectTitle", projectTitle || "Technical & Commercial Proposal");

      setState("converting");

      const response = await fetch("/api/convert", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.error || "Conversion failed");
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const name = `IOTTIVE_${(clientName || "Client").replace(/\s+/g, "_")}_Branded.docx`;

      setDownloadUrl(url);
      setDownloadName(name);
      setState("done");

      toast({
        title: "Conversion complete",
        description: "Your branded document is ready to download.",
      });
    } catch (error: any) {
      setState("error");
      setErrorMsg(error.message || "Something went wrong");
      toast({
        title: "Conversion failed",
        description: error.message,
        variant: "destructive",
      });
    }
  };

  const handleDownload = () => {
    if (!downloadUrl) return;
    const a = document.createElement("a");
    a.href = downloadUrl;
    a.download = downloadName;
    a.click();
  };

  const formatFileSize = (bytes: number) => {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  };

  return (
    <div className="min-h-screen bg-background flex flex-col">
      {/* Header */}
      <header className="border-b border-border px-6 py-4">
        <div className="max-w-3xl mx-auto flex items-center gap-3">
          <div className="w-8 h-8 bg-foreground rounded-full flex items-center justify-center">
            <span className="text-background text-xs font-bold">i</span>
          </div>
          <div>
            <h1 className="text-sm font-semibold tracking-tight" data-testid="text-app-title">
              IOTTIVE Doc Converter
            </h1>
            <p className="text-xs text-muted-foreground">
              Convert any document to branded format
            </p>
          </div>
        </div>
      </header>

      {/* Main */}
      <main className="flex-1 px-6 py-8">
        <div className="max-w-3xl mx-auto space-y-6">

          {/* Upload Zone */}
          <Card className="border border-border">
            <CardContent className="p-6">
              <div className="space-y-4">
                <div className="flex items-center justify-between">
                  <h2 className="text-sm font-semibold">Upload Document</h2>
                  <span className="text-xs text-muted-foreground">PDF or DOCX, max 20 MB</span>
                </div>

                {!file ? (
                  <div
                    data-testid="dropzone"
                    onDrop={handleDrop}
                    onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                    onDragLeave={() => setDragOver(false)}
                    onClick={() => fileInputRef.current?.click()}
                    className={`
                      border-2 border-dashed rounded-lg p-10 text-center cursor-pointer
                      transition-colors duration-150
                      ${dragOver
                        ? "border-primary bg-primary/5"
                        : "border-border hover:border-muted-foreground/40"
                      }
                    `}
                  >
                    <Upload className="w-8 h-8 mx-auto mb-3 text-muted-foreground" />
                    <p className="text-sm text-muted-foreground">
                      Drop your file here or <span className="text-primary font-medium">browse</span>
                    </p>
                    <input
                      ref={fileInputRef}
                      type="file"
                      accept=".pdf,.docx"
                      onChange={handleFileSelect}
                      className="hidden"
                      data-testid="input-file"
                    />
                  </div>
                ) : (
                  <div className="flex items-center gap-3 p-3 rounded-lg bg-muted/50 border border-border">
                    <FileText className="w-8 h-8 text-primary shrink-0" />
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium truncate" data-testid="text-filename">{file.name}</p>
                      <p className="text-xs text-muted-foreground">{formatFileSize(file.size)}</p>
                    </div>
                    {state === "done" && (
                      <CheckCircle2 className="w-5 h-5 text-green-500 shrink-0" />
                    )}
                    <button
                      onClick={(e) => { e.stopPropagation(); removeFile(); }}
                      className="p-1 rounded hover:bg-muted"
                      data-testid="button-remove-file"
                    >
                      <X className="w-4 h-4 text-muted-foreground" />
                    </button>
                  </div>
                )}
              </div>
            </CardContent>
          </Card>

          {/* Cover Page Details */}
          <Card className="border border-border">
            <CardContent className="p-6">
              <h2 className="text-sm font-semibold mb-4">Cover Page Details</h2>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <Label htmlFor="clientName" className="text-xs">Client / Company Name</Label>
                  <Input
                    id="clientName"
                    placeholder="e.g. Acme Corp"
                    value={clientName}
                    onChange={(e) => setClientName(e.target.value)}
                    data-testid="input-client-name"
                  />
                </div>
                <div className="space-y-1.5">
                  <Label htmlFor="contactName" className="text-xs">Contact Person</Label>
                  <Input
                    id="contactName"
                    placeholder="e.g. John Smith"
                    value={contactName}
                    onChange={(e) => setContactName(e.target.value)}
                    data-testid="input-contact-name"
                  />
                </div>
                <div className="space-y-1.5">
                  <Label htmlFor="contactEmail" className="text-xs">Contact Email</Label>
                  <Input
                    id="contactEmail"
                    type="email"
                    placeholder="e.g. john@acme.com"
                    value={contactEmail}
                    onChange={(e) => setContactEmail(e.target.value)}
                    data-testid="input-contact-email"
                  />
                </div>
                <div className="space-y-1.5">
                  <Label htmlFor="projectTitle" className="text-xs">Project Title</Label>
                  <Input
                    id="projectTitle"
                    placeholder="e.g. Mobile App Development"
                    value={projectTitle}
                    onChange={(e) => setProjectTitle(e.target.value)}
                    data-testid="input-project-title"
                  />
                </div>
              </div>
            </CardContent>
          </Card>

          {/* Action Buttons */}
          <div className="flex items-center gap-3">
            <Button
              onClick={handleConvert}
              disabled={!file || state === "uploading" || state === "converting"}
              className="flex-1"
              data-testid="button-convert"
            >
              {(state === "uploading" || state === "converting") ? (
                <>
                  <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                  {state === "uploading" ? "Uploading..." : "Converting..."}
                </>
              ) : (
                <>
                  <FileText className="w-4 h-4 mr-2" />
                  Convert to Branded DOCX
                </>
              )}
            </Button>

            {state === "done" && downloadUrl && (
              <Button
                onClick={handleDownload}
                variant="secondary"
                className="shrink-0"
                data-testid="button-download"
              >
                <Download className="w-4 h-4 mr-2" />
                Download
              </Button>
            )}
          </div>

          {state === "error" && errorMsg && (
            <p className="text-sm text-destructive" data-testid="text-error">{errorMsg}</p>
          )}

          {/* Brand Info */}
          <div className="rounded-lg border border-border p-4 space-y-2">
            <h3 className="text-xs font-semibold text-muted-foreground uppercase tracking-wide">
              What this does
            </h3>
            <ul className="text-xs text-muted-foreground space-y-1.5">
              <li className="flex items-start gap-2">
                <span className="text-primary mt-0.5">1.</span>
                Keeps the IOTTIVE branded cover page (logo, gradient, layout)
              </li>
              <li className="flex items-start gap-2">
                <span className="text-primary mt-0.5">2.</span>
                Updates cover with your client name, contact, and project title
              </li>
              <li className="flex items-start gap-2">
                <span className="text-primary mt-0.5">3.</span>
                Re-formats all content with Lexend font family and IOTTIVE styles
              </li>
              <li className="flex items-start gap-2">
                <span className="text-primary mt-0.5">4.</span>
                Preserves headings, bullet points, tables, and text formatting
              </li>
              <li className="flex items-start gap-2">
                <span className="text-primary mt-0.5">5.</span>
                Keeps headers and footers from the brand template
              </li>
            </ul>
          </div>

        </div>
      </main>

      {/* Footer */}
      <footer className="border-t border-border px-6 py-3">
        <p className="text-xs text-muted-foreground text-center">
          IOTTIVE Document Branding Tool — Internal Use
        </p>
      </footer>
    </div>
  );
}
