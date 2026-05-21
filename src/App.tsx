import React, { useState, useEffect, useCallback } from 'react';
import { 
  FileText, 
  CheckSquare, 
  ShieldAlert, 
  Send, 
  Download, 
  Loader2, 
  Plus, 
  FileJson,
  Wand2,
  Settings as SettingsIcon,
  Layout,
  Save,
  Cloud,
  ExternalLink,
  CheckCircle2,
  AlertCircle
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { generateFormStructure, ReferenceFilePayload } from './services/codexGeneratorService';
import { downloadDocx } from './services/docxService';
import { GeneratedForm, AppSettings } from './types';
import { cn } from './lib/utils';

const FORM_TYPES = [
  { id: 'test_sheet', label: 'Test Sheet', icon: FileText, color: 'text-blue-600', bg: 'bg-blue-50' },
  { id: 'checklist', label: 'Checklist', icon: CheckSquare, color: 'text-green-600', bg: 'bg-green-50' },
  { id: 'swms', label: 'SWMS', icon: ShieldAlert, color: 'text-orange-600', bg: 'bg-orange-50' },
  { id: 'other', label: 'Custom Form', icon: Plus, color: 'text-purple-600', bg: 'bg-purple-50' },
];

const DEFAULT_SETTINGS: AppSettings = {
  companyName: 'My Business Name',
  headerText: 'Safety & Compliance Document',
  footerText: 'Confidential - For Internal Use Only',
  showTable: true,
};

export default function App() {
  const [activeTab, setActiveTab] = useState<'builder' | 'settings'>('builder');
  const [selectedType, setSelectedType] = useState(FORM_TYPES[0].id);
  const [prompt, setPrompt] = useState('');
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isPushing, setIsPushing] = useState(false);
  const [isServiceM8Connected, setIsServiceM8Connected] = useState(false);
  const [pushSuccess, setPushSuccess] = useState<string | null>(null);
  const [result, setResult] = useState<GeneratedForm | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [settings, setSettings] = useState<AppSettings>(() => {
    const saved = localStorage.getItem('formai_settings');
    return saved ? JSON.parse(saved) : DEFAULT_SETTINGS;
  });

  const checkAuthStatus = useCallback(async () => {
    try {
      const response = await fetch('/api/auth/status');
      const data = await response.json();
      setIsServiceM8Connected(data.authenticated);
    } catch (err) {
      console.error('Failed to check auth status:', err);
    }
  }, []);

  useEffect(() => {
    checkAuthStatus();
  }, [checkAuthStatus]);

  useEffect(() => {
    const handleMessage = (event: MessageEvent) => {
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        checkAuthStatus();
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, [checkAuthStatus]);

  const handleConnectServiceM8 = async () => {
    try {
      const response = await fetch('/api/auth/url');
      const { url } = await response.json();
      window.open(url, 'sm8_oauth', 'width=600,height=700');
    } catch (err) {
      console.error('Failed to get auth URL:', err);
      setError('Failed to initiate ServiceM8 connection.');
    }
  };

  const handlePushToServiceM8 = async () => {
    if (!result || (!isServiceM8Connected && !settings.servicem8ApiKey)) return;

    setIsPushing(true);
    setError(null);
    setPushSuccess(null);

    try {
      const response = await fetch('/api/servicem8/push-form', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          form: result,
          apiKey: settings.servicem8ApiKey 
        }),
      });

      const data = await response.json();
      if (data.success) {
        setPushSuccess(`Successfully pushed to ServiceM8! Form UUID: ${data.formUuid}`);
      } else {
        const detailMsg = typeof data.details === 'object' ? JSON.stringify(data.details) : data.details;
        console.error('ServiceM8 Push Error Details:', data.details);
        throw new Error(`${data.error || 'Failed to push form'}${detailMsg ? `: ${detailMsg}` : ''}`);
      }
    } catch (err: any) {
      console.error(err);
      setError(err.message || 'Failed to push form to ServiceM8.');
    } finally {
      setIsPushing(false);
    }
  };

  const handleSaveSettings = (e: React.FormEvent) => {
    e.preventDefault();
    localStorage.setItem('formai_settings', JSON.stringify(settings));
    setActiveTab('builder');
  };

  const handleGenerate = async () => {
    if (!prompt.trim() && !selectedFile) {
      setError('Please provide requirements or upload a reference file.');
      return;
    }
    
    setIsGenerating(true);
    setError(null);
    try {
      let referenceFile: ReferenceFilePayload | undefined;
      
      if (selectedFile) {
        const reader = new FileReader();
        const base64Promise = new Promise<string>((resolve, reject) => {
          reader.onload = () => {
            const base64 = (reader.result as string).split(',')[1];
            resolve(base64);
          };
          reader.onerror = reject;
        });
        reader.readAsDataURL(selectedFile);
        const base64 = await base64Promise;
        referenceFile = {
          data: base64,
          mimeType: selectedFile.type || inferMimeType(selectedFile.name),
          name: selectedFile.name,
        };
      }

      const typeLabel = FORM_TYPES.find(t => t.id === selectedType)?.label || selectedType;
      const data = await generateFormStructure(prompt || "Generate a form based on the attached reference file.", typeLabel, referenceFile);
      
      // Sanitize field labels for ServiceM8 compatibility
      const sanitizedData = {
        ...data,
        fields: data.fields.map(field => ({
          ...field,
          label: field.label
            .replace(/[^a-zA-Z0-9\s]/g, " ")
            .replace(/\s+/g, " ")
            .trim()
        }))
      };
      
      setResult(sanitizedData);
    } catch (err: any) {
      console.error(err);
      setError(err.message || 'Failed to generate form. Please try again.');
    } finally {
      setIsGenerating(false);
    }
  };

  const handleDownloadDocx = async () => {
    if (!result) return;
    await downloadDocx(result, settings);
  };

  const sanitizeForMergeField = (label: string) => {
    return label
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_|_$/g, "");
  };

  const inferMimeType = (fileName: string) => {
    const extension = fileName.split('.').pop()?.toLowerCase();
    const mimeTypes: Record<string, string> = {
      pdf: 'application/pdf',
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      webp: 'image/webp',
      txt: 'text/plain',
      md: 'text/markdown',
      csv: 'text/csv',
      docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    };

    return extension ? (mimeTypes[extension] || 'application/octet-stream') : 'application/octet-stream';
  };

  const handleDownloadJson = () => {
    if (!result) return;
    
    const sm8Payload = {
      form: {
        name: result.title,
        badge_name: result.badgeName.substring(0, 11),
        description: result.description,
        active: 1
      },
      fields: result.fields.map((field, index) => {
        const typeMap: Record<string, string> = {
          text: "Text",
          number: "Number",
          date: "Date",
          checkbox: "Multiple Choice",
          select: "Multiple Choice",
          signature: "Signature",
          photo: "Photo"
        };
        
        const fieldData: any = {
          fieldType: typeMap[field.type] || "Text",
          additionalDetails: field.additionalDetails || "",
          mandatory: !!field.required,
          conditions: [
            { question: "", operator: "", value: "" },
            { question: "", operator: "", value: "" },
            { question: "", operator: "", value: "" }
          ],
          conditionMethod: "AND"
        };

        if (field.type === 'select' || field.type === 'checkbox') {
          fieldData.choices = field.options && field.options.length > 0 
            ? field.options 
            : ["Yes", "No"];
        }
        
        return {
          name: field.label,
          field_data_json: JSON.stringify(fieldData),
          sort_order: String(index + 1)
        };
      })
    };

    const blob = new Blob([JSON.stringify(sm8Payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${result.title.replace(/\s+/g, '_')}_servicem8.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 px-6 py-4 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center text-white shadow-lg shadow-indigo-100">
              <Wand2 size={24} />
            </div>
            <div>
              <h1 className="text-xl font-bold text-slate-900 tracking-tight">FormAI Builder</h1>
              <p className="text-xs text-slate-500 font-medium uppercase tracking-wider">ServiceM8 & DOCX Generator</p>
            </div>
          </div>

          <nav className="flex bg-slate-100 p-1 rounded-xl">
            <button
              onClick={() => setActiveTab('builder')}
              className={cn(
                "flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-all",
                activeTab === 'builder' ? "bg-white text-indigo-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
              )}
            >
              <Layout size={18} />
              Builder
            </button>
            <button
              onClick={() => setActiveTab('settings')}
              className={cn(
                "flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-all",
                activeTab === 'settings' ? "bg-white text-indigo-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
              )}
            >
              <SettingsIcon size={18} />
              Settings
            </button>
          </nav>
        </div>
      </header>

      <main className="flex-1 max-w-6xl mx-auto w-full p-6">
        <AnimatePresence mode="wait">
          {activeTab === 'builder' ? (
            <motion.div
              key="builder"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="grid grid-cols-1 lg:grid-cols-12 gap-8"
            >
              {/* Left Column: Input */}
              <div className="lg:col-span-5 space-y-6">
                <section className="bg-white rounded-2xl p-6 shadow-sm border border-slate-200">
                  <h2 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4">1. Select Form Type</h2>
                  <div className="grid grid-cols-2 gap-3">
                    {FORM_TYPES.map((type) => (
                      <button
                        key={type.id}
                        onClick={() => setSelectedType(type.id)}
                        className={cn(
                          "flex flex-col items-center justify-center p-4 rounded-xl border-2 transition-all duration-200",
                          selectedType === type.id 
                            ? "border-indigo-600 bg-indigo-50/50" 
                            : "border-slate-100 hover:border-slate-200 bg-white"
                        )}
                      >
                        <type.icon className={cn("mb-2", type.color)} size={24} />
                        <span className="text-sm font-medium text-slate-700">{type.label}</span>
                      </button>
                    ))}
                  </div>
                </section>

                <section className="bg-white rounded-2xl p-6 shadow-sm border border-slate-200">
                  <h2 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4">2. Describe your requirements</h2>
                  <div className="space-y-4">
                    <textarea
                      value={prompt}
                      onChange={(e) => setPrompt(e.target.value)}
                      placeholder="e.g. A safety checklist for high-voltage electrical work including PPE verification, area clearance, and supervisor sign-off..."
                      className="w-full h-48 p-4 rounded-xl border border-slate-200 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all resize-none text-slate-700 leading-relaxed"
                    />

                    <div className="relative">
                      <label className="block text-xs font-bold text-slate-400 uppercase tracking-widest mb-2">
                        Reference File
                      </label>
                      <div 
                        onClick={() => document.getElementById('pdf-upload')?.click()}
                        className={cn(
                          "flex flex-col items-center justify-center w-full p-6 border-2 border-dashed rounded-xl transition-all cursor-pointer",
                          selectedFile ? "border-indigo-500 bg-indigo-50" : "border-slate-200 hover:border-indigo-400 bg-slate-50"
                        )}
                      >
                        <input
                          id="pdf-upload"
                          type="file"
                          accept=".pdf,.png,.jpg,.jpeg,.webp,.txt,.md,.csv,.docx,application/pdf,image/png,image/jpeg,image/webp,text/plain,text/markdown,text/csv,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                          className="hidden"
                          onChange={(e) => {
                            const file = e.target.files?.[0];
                            if (file) setSelectedFile(file);
                          }}
                        />
                        <FileText className={cn("w-8 h-8 mb-2", selectedFile ? "text-indigo-600" : "text-slate-400")} />
                        <p className="text-sm text-slate-600 text-center">
                          {selectedFile ? selectedFile.name : "Upload a PDF, image, text file, CSV, or DOCX"}
                        </p>
                        {selectedFile && (
                          <button 
                            onClick={(e) => {
                              e.stopPropagation();
                              setSelectedFile(null);
                            }}
                            className="mt-2 text-xs text-red-500 hover:text-red-700 font-medium"
                          >
                            Remove File
                          </button>
                        )}
                      </div>
                    </div>
                  </div>
                  <button
                    onClick={handleGenerate}
                    disabled={isGenerating || (!prompt.trim() && !selectedFile)}
                    className="w-full mt-4 bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-300 text-white font-semibold py-3 px-6 rounded-xl flex items-center justify-center gap-2 transition-all shadow-lg shadow-indigo-200"
                  >
                    {isGenerating ? (
                      <>
                        <Loader2 className="animate-spin" size={20} />
                        <span>Generating...</span>
                      </>
                    ) : (
                      <>
                        <Send size={20} />
                        <span>Generate Form & Template</span>
                      </>
                    )}
                  </button>
                  {error && <p className="mt-3 text-sm text-red-500 text-center">{error}</p>}
                </section>
              </div>

              {/* Right Column: Output */}
              <div className="lg:col-span-7">
                <AnimatePresence mode="wait">
                  {!result ? (
                    <motion.div
                      initial={{ opacity: 0, y: 20 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: -20 }}
                      className="h-full flex flex-col items-center justify-center p-12 bg-white rounded-2xl border-2 border-dashed border-slate-200 text-slate-400"
                    >
                      <div className="w-16 h-16 bg-slate-50 rounded-full flex items-center justify-center mb-4">
                        <FileText size={32} />
                      </div>
                      <p className="text-lg font-medium">Your generated form will appear here</p>
                      <p className="text-sm">Describe what you need and click generate</p>
                    </motion.div>
                  ) : (
                    <motion.div
                      initial={{ opacity: 0, scale: 0.95 }}
                      animate={{ opacity: 1, scale: 1 }}
                      className="space-y-6"
                    >
                      {/* Result Header */}
                      <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-200">
                        <div className="flex items-start justify-between mb-6">
                          <div className="flex-1 mr-4">
                            <h2 className="text-2xl font-bold text-slate-900">{result.title}</h2>
                            <div className="flex items-center gap-2 mt-1">
                              <span className="text-[10px] font-bold bg-indigo-50 text-indigo-600 px-2 py-0.5 rounded uppercase tracking-wider border border-indigo-100">
                                Badge: {result.badgeName}
                              </span>
                              <p className="text-slate-500 text-sm">{result.description}</p>
                            </div>
                          </div>
          <div className="flex gap-2 shrink-0">
            <button
              onClick={handleDownloadDocx}
              className="p-2 bg-blue-50 text-blue-600 rounded-lg hover:bg-blue-100 transition-colors"
              title="Download DOCX Template"
            >
              <Download size={20} />
            </button>
            <button
              onClick={handleDownloadJson}
              className="p-2 bg-indigo-50 text-indigo-600 rounded-lg hover:bg-indigo-100 transition-colors"
              title="Download ServiceM8 JSON"
            >
              <FileJson size={20} />
            </button>
            {(isServiceM8Connected || settings.servicem8ApiKey) ? (
              <button
                onClick={handlePushToServiceM8}
                disabled={isPushing}
                className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors disabled:bg-slate-300 shadow-sm"
                title="Push to ServiceM8"
              >
                {isPushing ? <Loader2 className="animate-spin" size={18} /> : <Cloud size={18} />}
                <span className="text-sm font-semibold">
                  {settings.servicem8ApiKey ? 'Push (API Key)' : 'Push to ServiceM8'}
                </span>
              </button>
            ) : (
              <button
                onClick={handleConnectServiceM8}
                className="flex items-center gap-2 px-4 py-2 bg-slate-800 text-white rounded-lg hover:bg-slate-900 transition-colors shadow-sm"
                title="Connect to ServiceM8"
              >
                <ExternalLink size={18} />
                <span className="text-sm font-semibold">Connect ServiceM8</span>
              </button>
            )}
          </div>
        </div>

        {pushSuccess && (
          <div className="mb-6 p-4 bg-green-50 border border-green-200 rounded-xl flex items-center gap-3 text-green-700">
            <CheckCircle2 size={20} className="shrink-0" />
            <p className="text-sm font-medium">{pushSuccess}</p>
          </div>
        )}

        {error && !isGenerating && !isPushing && (
          <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-xl flex items-center gap-3 text-red-700">
            <AlertCircle size={20} className="shrink-0" />
            <p className="text-sm font-medium">{error}</p>
          </div>
        )}

                        {/* Tabs/Sections */}
                        <div className="space-y-8">
                          {/* ServiceM8 Fields */}
                          <div>
                            <h3 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4 flex items-center gap-2">
                              <CheckSquare size={14} />
                              ServiceM8 Form Fields
                            </h3>
                            <div className="space-y-2">
                              {result.fields.map((field, idx) => (
                                <div key={idx} className="flex items-center justify-between p-3 bg-slate-50 rounded-xl border border-slate-100">
                                  <div className="flex items-center gap-3">
                                    <div className="w-8 h-8 bg-white rounded-lg border border-slate-200 flex items-center justify-center text-slate-400">
                                      <span className="text-[10px] font-bold uppercase">{field.type.slice(0, 3)}</span>
                                    </div>
                                    <div>
                                      <p className="text-sm font-semibold text-slate-700">{field.label}</p>
                                      {field.additionalDetails && (
                                        <p className="text-[10px] text-slate-500 italic mt-0.5">{field.additionalDetails}</p>
                                      )}
                                      <p className="text-[10px] text-slate-400 font-mono mt-1">
                                        Native Merge Field: MERGEFIELD {field.type === 'photo' ? 'image_' : ''}form_{sanitizeForMergeField(field.label)}{field.type === 'photo' ? '_medium' : ''} \* MERGEFORMAT
                                      </p>
                                    </div>
                                  </div>
                                  {field.required && (
                                    <span className="text-[10px] bg-red-50 text-red-500 px-2 py-0.5 rounded-full font-bold uppercase">Required</span>
                                  )}
                                </div>
                              ))}
                            </div>
                          </div>

                          {/* DOCX Preview */}
                          <div>
                            <h3 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4 flex items-center gap-2">
                              <FileText size={14} />
                              DOCX Template Preview
                            </h3>
                            <div className="bg-white rounded-xl border border-slate-200 shadow-inner overflow-hidden">
                              <div className="bg-slate-50 border-b border-slate-200 px-4 py-2 flex items-center justify-between">
                                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Document Layout</span>
                                <div className="flex gap-1">
                                  <div className="w-2 h-2 rounded-full bg-red-400" />
                                  <div className="w-2 h-2 rounded-full bg-amber-400" />
                                  <div className="w-2 h-2 rounded-full bg-green-400" />
                                </div>
                              </div>
                              <div className="p-8 space-y-6 max-h-[600px] overflow-y-auto bg-white">
                                <div className="text-center mb-8">
                                  <h4 className="text-xl font-bold text-slate-900">{result.title}</h4>
                                  <p className="text-xs text-slate-500 italic mt-1">{result.description}</p>
                                </div>

                                {result.docxContent.sections.map((section, idx) => (
                                  <div key={idx} className={cn(
                                    "space-y-2",
                                    section.isStandardHeader && "bg-slate-50 p-4 rounded-lg border border-slate-100",
                                    section.isStandardFooter && "border-t border-slate-100 pt-4 text-right"
                                  )}>
                                    {!section.isStandardHeader && !section.isStandardFooter && (
                                      <h5 className="text-sm font-bold text-indigo-600 border-b border-indigo-50 pb-1">{section.title}</h5>
                                    )}
                                    
                                    {section.layout === 'table' || section.isStandardHeader ? (
                                      <div className="grid grid-cols-1 gap-1">
                                        {section.content.split('\n').filter(l => l.trim()).map((line, lIdx) => {
                                          const [label, ...rest] = line.split(':');
                                          const value = rest.join(':').trim();
                                          
                                          // Helper to format preview text (handle IF statements)
                                          const formatPreviewValue = (val: string) => {
                                            if (val.includes('{IF')) {
                                              // Extract symbols from {IF "..."="Yes" "☑" "☐"}
                                              const match = val.match(/"([^"]+)"\s+"([^"]+)"\s*\}$/);
                                              if (match) {
                                                return (
                                                  <span className="flex items-center gap-1">
                                                    <span className="text-indigo-600 font-bold">{match[1]}</span>
                                                    <span className="text-slate-300">/</span>
                                                    <span>{match[2]}</span>
                                                    <span className="ml-2 text-[9px] text-slate-400 font-mono">(IF Logic)</span>
                                                  </span>
                                                );
                                              }
                                            }
                                            return val;
                                          };

                                          return (
                                            <div key={lIdx} className="grid grid-cols-3 gap-2 text-[11px] border-b border-slate-50 pb-1 last:border-0">
                                              <span className="font-bold text-slate-600">{label.trim()}:</span>
                                              <span className="col-span-2 text-slate-500 font-mono bg-slate-50/50 px-1 rounded">
                                                {formatPreviewValue(value || line)}
                                              </span>
                                            </div>
                                          );
                                        })}
                                      </div>
                                    ) : (
                                      <div className="text-[11px] text-slate-600 leading-relaxed whitespace-pre-wrap">
                                        {section.content.split('\n').map((line, lIdx) => (
                                          <div key={lIdx} className="mb-1">
                                            {line.includes('{IF') ? (
                                              <span className="font-mono bg-slate-50/50 px-1 rounded">
                                                {/* Re-using the logic for inline IF statements */}
                                                {(() => {
                                                  const match = line.match(/"([^"]+)"\s+"([^"]+)"\s*\}$/);
                                                  if (match) {
                                                    const prefix = line.split('{IF')[0];
                                                    return (
                                                      <span className="flex items-center gap-1">
                                                        <span>{prefix}</span>
                                                        <span className="text-indigo-600 font-bold">{match[1]}</span>
                                                        <span className="text-slate-300">/</span>
                                                        <span>{match[2]}</span>
                                                      </span>
                                                    );
                                                  }
                                                  return line;
                                                })()}
                                              </span>
                                            ) : line}
                                          </div>
                                        ))}
                                      </div>
                                    )}
                                  </div>
                                ))}
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>
            </motion.div>
          ) : (
            <motion.div
              key="settings"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="max-w-2xl mx-auto"
            >
              <div className="bg-white rounded-2xl p-8 shadow-sm border border-slate-200">
                <div className="flex items-center gap-3 mb-8">
                  <div className="w-12 h-12 bg-slate-100 rounded-xl flex items-center justify-center text-slate-600">
                    <SettingsIcon size={24} />
                  </div>
                  <div>
                    <h2 className="text-2xl font-bold text-slate-900">Document Settings</h2>
                    <p className="text-slate-500">Customize the appearance of your generated DOCX files.</p>
                  </div>
                </div>

                <form onSubmit={handleSaveSettings} className="space-y-6">
                  <div className="space-y-2">
                    <label className="text-sm font-semibold text-slate-700">Company Name</label>
                    <input
                      type="text"
                      value={settings.companyName}
                      onChange={(e) => setSettings({ ...settings, companyName: e.target.value })}
                      className="w-full p-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-indigo-500 outline-none transition-all"
                      placeholder="e.g. Acme Services Pty Ltd"
                    />
                  </div>

                  <div className="space-y-2">
                    <label className="text-sm font-semibold text-slate-700">Header Text</label>
                    <textarea
                      value={settings.headerText}
                      onChange={(e) => setSettings({ ...settings, headerText: e.target.value })}
                      className="w-full p-3 h-24 rounded-xl border border-slate-200 focus:ring-2 focus:ring-indigo-500 outline-none transition-all resize-none"
                      placeholder="Appears at the top of every page..."
                    />
                  </div>

                  <div className="space-y-2">
                    <label className="text-sm font-semibold text-slate-700">Footer Text</label>
                    <textarea
                      value={settings.footerText}
                      onChange={(e) => setSettings({ ...settings, footerText: e.target.value })}
                      className="w-full p-3 h-24 rounded-xl border border-slate-200 focus:ring-2 focus:ring-indigo-500 outline-none transition-all resize-none"
                      placeholder="Appears at the bottom of every page..."
                    />
                  </div>

                  <div className="space-y-2">
                    <label className="text-sm font-semibold text-slate-700">ServiceM8 API Key (for testing)</label>
                    <input
                      type="password"
                      value={settings.servicem8ApiKey || ''}
                      onChange={(e) => setSettings({ ...settings, servicem8ApiKey: e.target.value })}
                      className="w-full p-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-indigo-500 outline-none transition-all"
                      placeholder="Enter your ServiceM8 API Key..."
                    />
                    <p className="text-[10px] text-slate-400">
                      Found in ServiceM8 &gt; Settings &gt; ServiceM8 Add-ons &gt; API. 
                      If provided, this will be used instead of OAuth.
                    </p>
                  </div>

                  <div className="flex items-center justify-between p-4 bg-slate-50 rounded-xl border border-slate-100">
                    <div>
                      <p className="text-sm font-semibold text-slate-700">Table Layout</p>
                      <p className="text-xs text-slate-500">Display form fields in a structured table</p>
                    </div>
                    <button
                      type="button"
                      onClick={() => setSettings({ ...settings, showTable: !settings.showTable })}
                      className={cn(
                        "w-12 h-6 rounded-full transition-all relative",
                        settings.showTable ? "bg-indigo-600" : "bg-slate-300"
                      )}
                    >
                      <div className={cn(
                        "absolute top-1 w-4 h-4 bg-white rounded-full transition-all",
                        settings.showTable ? "left-7" : "left-1"
                      )} />
                    </button>
                  </div>

                  <button
                    type="submit"
                    className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-4 px-6 rounded-xl flex items-center justify-center gap-2 transition-all shadow-lg shadow-indigo-200"
                  >
                    <Save size={20} />
                    Save Settings
                  </button>
                </form>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Footer */}
      <footer className="bg-white border-t border-slate-200 py-8 px-6 mt-auto">
        <div className="max-w-6xl mx-auto flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-2">
            <div className="w-6 h-6 bg-indigo-100 rounded flex items-center justify-center text-indigo-600">
              <Wand2 size={14} />
            </div>
            <p className="text-sm text-slate-500 font-medium">
              &copy; 2026 FormAI Builder.
            </p>
          </div>
          <div className="flex items-center gap-8">
            <a href="#" className="text-sm text-slate-400 hover:text-indigo-600 transition-colors">Documentation</a>
            <a href="#" className="text-sm text-slate-400 hover:text-indigo-600 transition-colors">ServiceM8 Help</a>
            <a href="#" className="text-sm text-slate-400 hover:text-indigo-600 transition-colors">Privacy Policy</a>
          </div>
        </div>
      </footer>
    </div>
  );
}
