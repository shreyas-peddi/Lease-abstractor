/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import ReactDOM from 'react-dom/client';
import { GoogleGenAI, Type } from '@google/genai';
import * as pdfjsLib from 'pdfjs-dist';
import * as XLSX from 'xlsx';
import Tesseract from 'tesseract.js';

// Set up the PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://aistudiocdn.com/pdfjs-dist@^4.5.136/build/pdf.worker.min.mjs`;

// Helper to convert camelCase to Title Case for display
const toTitleCase = (str: string) => {
  if (!str) return '';
  return str.replace(/([A-Z])/g, ' $1').replace(/^./, (s) => s.toUpperCase());
};

const sectionIcons: { [key: string]: string } = {
  generalInformation: 'fa-solid fa-building-user',
  leaseAmendmentsReviewed: 'fa-solid fa-file-signature',
  leaseTermAndDates: 'fa-solid fa-calendar-days',
  noticeAddresses: 'fa-solid fa-map-location-dot',
  billingAndCharges: 'fa-solid fa-file-invoice-dollar',
  leaseNotes: 'fa-solid fa-clipboard-list',
  maintenanceAndReimbursement: 'fa-solid fa-wrench',
  tenantInsuranceInformation: 'fa-solid fa-shield-halved',
};

// FIX: Add explicit return type React.ReactElement to avoid JSX namespace issues.
const App = (): React.ReactElement => {
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [abstractData, setAbstractData] = useState<any | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [progressMessage, setProgressMessage] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement>(null);
  const leaseTextCache = useRef<string>('');
  const chatHistoryRef = useRef<HTMLDivElement>(null);

  // Q&A State
  const [userQuestion, setUserQuestion] = useState<string>('');
  const [chatHistory, setChatHistory] = useState<{ role: 'user' | 'model'; content: string }[]>([]);
  const [isAnswering, setIsAnswering] = useState(false);
  const [qaError, setQaError] = useState<string | null>(null);

  // Auto-scroll chat history
  useEffect(() => {
    if (chatHistoryRef.current) {
      chatHistoryRef.current.scrollTop = chatHistoryRef.current.scrollHeight;
    }
  }, [chatHistory, isAnswering]);

const addressSchema = {
  type: Type.OBJECT,
  properties: {
    name: { type: Type.STRING },
    atten: { type: Type.STRING },
    address: { type: Type.STRING },
    city: { type: Type.STRING },
    state: { type: Type.STRING },
    zipCode: { type: Type.STRING },
    email: { type: Type.STRING },
  },
};

const clauseSchema = {
  type: Type.OBJECT,
  properties: {
    leaseSectionAndPage: { type: Type.STRING },
    notes: { type: Type.STRING },
  },
};

const leaseAbstractSchema = {
  type: Type.OBJECT,
  properties: {
    generalInformation: {
      type: Type.OBJECT,
      properties: {
        tenantName: { type: Type.STRING },
        dba: { type: Type.STRING },
        suiteNumber: { type: Type.STRING },
        buildingNumber: { type: Type.STRING },
        storeNumber: { type: Type.STRING },
        shoppingCenterName: { type: Type.STRING },
        shoppingCenterAddress: { type: Type.STRING },
        premisesGLA: { type: Type.STRING },
        guarantors: { type: Type.STRING },
        areSpousesGuarantors: { type: Type.STRING },
        isGuaranteeSeparate: { type: Type.STRING },
      },
    },

    leaseAmendmentsReviewed: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          dateOfDocument: { type: Type.STRING },
          documentReviewed: { type: Type.STRING },
          leaseSectionAndPage: { type: Type.STRING },
          issues: { type: Type.STRING,description:"lists any differences or changes from previous version of lease" },
        },
      },
    },

    leaseTermAndDates: {
      type: Type.OBJECT,
      properties: {
        leaseTerm: { type: Type.STRING },
        leaseSectionAndPage: { type: Type.STRING },
        leaseExecutionDate: { type: Type.STRING },
        deliveryDate: { type: Type.STRING },
        leaseCommencementDate: { type: Type.STRING },
        openDate: { type: Type.STRING },
        rentCommencementDate: { type: Type.STRING },
        leaseExpirationDate: { type: Type.STRING },
      },
    },

    noticeAddresses: {
      type: Type.OBJECT,
      properties: {
        leaseSectionAndPage: { type: Type.STRING },
        tenant: addressSchema,
        tenantsLawyer: addressSchema,
        franchisor: {
          type: Type.OBJECT,
          properties: {
            ...addressSchema.properties,
            leaseAddressType: { type: Type.STRING },
          },
        },
      },
    },

    billingAndCharges: {
      type: Type.OBJECT,
      properties: {
        baseRentSchedule: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              incomeCategory: { type: Type.STRING,description:"It can be Free Rent or RNT" },
              effectiveDate: { type: Type.STRING },
              endDate: { type: Type.STRING },
              annualAmountPerSf: { type: Type.STRING },
              annualTotal: { type: Type.STRING },
              monthlyAmount: { type: Type.STRING },
              AreaInsf:{type:Type.STRING},
              leaseSectionAndPage: { type: Type.STRING },
            },
          },
        },
        camTaxInsuranceFirstYear: { type: Type.STRING },
        percentageRent: {
          type: Type.OBJECT,
          properties: {
            leaseSectionAndPage: { type: Type.STRING },
            reportingFrequency: { type: Type.STRING },
            naturalBreakpoint: { type: Type.STRING },
            unnaturalBreakpoint: { type: Type.STRING },
            salesYearEnd: { type: Type.STRING },
            billingFrequency: { type: Type.STRING },
          },
        },
      },
    },

        leaseNotes: {
  type: Type.OBJECT,
  properties: {
    securityDeposit: {
      ...clauseSchema,
      description: "Amount held by landlord as security for tenant obligations."
    },
    prepaidRent: {
      ...clauseSchema,
      description: "First month's rent or advanced rent paid before occupancy."
    },
    primaryUse: {
      ...clauseSchema,
      description: "Specifies the main permitted use of the premises by the tenant."
    },
    radiusRestrictions: {
      ...clauseSchema,
      description: "Limits tenant from opening another store within a certain distance from the shopping center."
    },
    exclusiveUse: {
      ...clauseSchema,
      description: "Tenant has exclusive rights to operate a specific business type; landlord cannot lease to competitors."
    },
    cotenancyRequirements: {
      ...clauseSchema,
      description: "Conditions about other tenants that must be open/operating for rent obligations or continued operation."
    },
    acceptableReplacementCotenancy: {
      ...clauseSchema,
      description: "Allows replacement of key tenants if original ones leave, maintaining co-tenancy requirements."
    },
    landlordKickout: {
      ...clauseSchema,
      description: "Landlord’s right to terminate the lease under certain conditions."
    },
    tenantKickout: {
      ...clauseSchema,
      description: "Tenant’s right to terminate the lease, in whole or part, under specified circumstances."
    },
    tenantGoDark: {
      ...clauseSchema,
      description: "Tenant’s right to close their store and implications for rent payment."
    },
    landlordRestrictions: {
      ...clauseSchema,
      description: "Limits on landlord’s ability to lease, develop, or modify the shopping center."
    },
    assignmentAndSubletting: {
      ...clauseSchema,
      description: "Tenant’s rights to assign or sublet their space."
    },
    shoppingCenterAlterations: {
      ...clauseSchema,
      description: "What changes the landlord can make to the shopping center, with or without tenant approval."
    },
    operatingCovenant: {
      ...clauseSchema,
      description: "Tenant’s obligation to open and continue operating at the shopping center."
    },
    lateChargesNSFFee: {
      ...clauseSchema,
      description: "Terms for handling late payments, insufficient funds, and defaults."
    },
    defaultClause: {
      ...clauseSchema,
      description: "Provisions outlining what constitutes a default and related remedies."
    },
    guaranty: {
      ...clauseSchema,
      description: "Guarantee provisions for lease obligations."
    },
    purchaseOptionROFR: {
      ...clauseSchema,
      description: "Tenant’s rights to purchase or make offers on the property."
    },
    marketingOrPromotionalFee: {
      ...clauseSchema,
      description: "Tenant’s obligation to contribute to marketing or promotional costs."
    },
    holdoverTerms: {
      ...clauseSchema,
      description: "Terms for tenant remaining after lease expiration."
    },
    signage: {
      ...clauseSchema,
      description: "Rules regarding tenant signage."
    },
    estoppel: {
      ...clauseSchema,
      description: "Documents that confirm key lease terms or amend them as needed."
    },
    eminentDomainAndSubordination: {
      ...clauseSchema,
      description: "Provisions for government taking (eminent domain) or lease subordination."
    },
    damageOrDestruction: {
      ...clauseSchema,
      description: "Terms for handling property damage or destruction."
    },
    relocationRight: {
      ...clauseSchema,
      description: "Landlord’s right to relocate tenant within the shopping center."
    },
    utilitiesHVACTrash: {
      ...clauseSchema,
      description: "Tenant’s responsibilities for utilities, HVAC, and waste removal."
    },
    repairsMaintenanceReplacements: {
      ...clauseSchema,
      description: "Responsibilities for repairs, maintenance, and replacements."
    },
    tenantImprovementAllowance: {
      ...clauseSchema,
      description: "Financial contributions from landlord for tenant improvements."
    },
    constructionManagementFees: {
      ...clauseSchema,
      description: "Terms for construction management, oversight, and general contractor fees."
    },
    landlordMaintenance: {
      ...clauseSchema,
      description: "Landlord’s responsibilities for maintenance of common areas and property."
    },
    brokersAndAgents: {
      ...clauseSchema,
      description: "Provisions regarding brokers and agents involved in the lease."
    },
    financialReporting: {
      ...clauseSchema,
      description: "Tenant’s obligation to provide financial reports."
    },
    parking: {
      ...clauseSchema,
      description: "Parking provisions for tenants, including allocation and shared use."
    },
    landlordWorkAndDelivery: {
      ...clauseSchema,
      description: "Landlord’s obligations before tenant takes possession and delivery condition."
    },
    tenantPlanSubmission: {
      ...clauseSchema,
      description: "Tenant’s obligation to submit plans before commencing work."
    },
    camTaxesInsurance: {
      ...clauseSchema,
      description: "Sections covering tenant’s obligations for common area maintenance, taxes, and insurance."
    },
    tenantDirectBill: {
      ...clauseSchema,
      description: "Costs paid directly by tenant, such as utilities."
    },
    Other: {
      type: Type.ARRAY,
      description: "Any additional responsibilities, obligations, or liabilities not covered above.",
      items: clauseSchema,
    },
  },
},

    maintenanceAndReimbursement: {
      type: Type.OBJECT,
      properties: {
        hvac: { type: Type.STRING },
        tenantAllowance: { type: Type.STRING },
        cam: {
          type: Type.OBJECT,
          properties: {
            leaseSectionAndPage: { type: Type.STRING },
            prorataSharePercent: { type: Type.STRING },
            exclusions: { type: Type.STRING },
            paymentTerms: { type: Type.STRING },
            capitalRepairs: { type: Type.STRING },
            auditRight: { type: Type.STRING },
            adminFeeAllowedInCAM: { type: Type.STRING },
            propertyManagementFeeAllowedInCAM: { type: Type.STRING },
          },
        },
        realEstateTaxes: {
          type: Type.OBJECT,
          properties: {
            leaseSectionAndPage: { type: Type.STRING },
            prorataSharePercent: { type: Type.STRING },
            paymentTerms: { type: Type.STRING },
            appealRight: { type: Type.STRING },
            tenantPaysForAssessments: { type: Type.STRING },
          },
        },
        insurance: {
          type: Type.OBJECT,
          properties: {
            leaseSectionAndPage: { type: Type.STRING },
            insuranceReimbursed: { type: Type.STRING },
            tenantPaysDeductible: { type: Type.STRING },
            prorataSharePercent: { type: Type.STRING },
            rightToSelfInsure: { type: Type.STRING },
          },
        },
      },
    },

    tenantInsuranceInformation: {
      type: Type.OBJECT,
      properties: {
        leaseSectionAndPage: { type: Type.STRING },
        coverages: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              type: { type: Type.STRING },
              coverage: { type: Type.STRING },
            },
          },
        },
        deductibleRequirementInLease: { type: Type.STRING },
        comments: { type: Type.STRING },
      },
    },
  },
};


  
  const handleFileChange = (files: FileList | null) => {
    if (!files) return;

    const newFiles = Array.from(files).filter(file => 
      file.type === 'application/pdf' && !selectedFiles.some(f => f.name === file.name)
    );
    
    if (newFiles.length > 0) {
      setSelectedFiles(prevFiles => [...prevFiles, ...newFiles].sort((a, b) => a.name.localeCompare(b.name)));
      setError(null);
    }
    
    const nonPdfFiles = Array.from(files).some(file => file.type !== 'application/pdf');
    if (nonPdfFiles) {
      setError('Only PDF files are accepted. Non-PDF files have been ignored.');
    }
  };

  const handleRemoveFile = (indexToRemove: number) => {
    const newFiles = selectedFiles.filter((_, index) => index !== indexToRemove);
    setSelectedFiles(newFiles);
    if (newFiles.length === 0) {
      setAbstractData(null); // Clear data if no files are left
      setChatHistory([]); // Clear chat
      leaseTextCache.current = ''; // Clear cache
    }
  };
  
  const extractTextFromPdf = async (file: File, onProgress: (message: string) => void): Promise<string> => {
    onProgress(`Reading file: ${file.name}`);
    const reader = new FileReader();
    const fileReadPromise = new Promise<ArrayBuffer>((resolve, reject) => {
        reader.onload = (event) => {
            if (!event.target?.result) {
                return reject(new Error("Failed to read file."));
            }
            resolve(event.target.result as ArrayBuffer);
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });

    const arrayBuffer = await fileReadPromise;
    const typedArray = new Uint8Array(arrayBuffer);
    const pdf = await pdfjsLib.getDocument(typedArray).promise;
    let fullText = '';
    let worker: Tesseract.Worker | null = null;

    try {
        for (let i = 1; i <= pdf.numPages; i++) {
            onProgress(`${file.name} - Page ${i}/${pdf.numPages}`);
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map(item => ('str' in item ? item.str : '')).join(' ');

            // Heuristic to detect image-based pages: few text items or very short total text length.
            if (textContent.items.length < 15 && pageText.trim().length < 150) {
                onProgress(`${file.name} - Page ${i}/${pdf.numPages} (OCR Scan)`);
                
                // Initialize worker only when first needed
                if (!worker) {
                    worker = await Tesseract.createWorker();
                }

                const viewport = page.getViewport({ scale: 2.0 }); // Higher scale for better OCR quality
                const canvas = document.createElement('canvas');
                const context = canvas.getContext('2d');
                if (!context) throw new Error("Could not get canvas context");
                canvas.height = viewport.height;
                canvas.width = viewport.width;

                const renderContext = {
                    canvasContext: context,
                    viewport: viewport
                };
                await page.render(renderContext as any).promise;
                
                const { data: { text } } = await worker.recognize(canvas);
                fullText += text + '\n\n';
            } else {
                fullText += pageText + '\n\n';
            }
            page.cleanup();
        }
    } finally {
        if (worker) {
            await worker.terminate();
        }
    }

    return fullText;
};

  const handleGenerateAbstract = async () => {
    if (selectedFiles.length === 0) {
      setError('Please upload one or more lease documents.');
      return;
    }
    setIsLoading(true);
    setError(null);
    setAbstractData(null);
    const totalFiles = selectedFiles.length;
    setProgressMessage(`Preparing ${totalFiles} document(s)...`);

    const fullAbstractData = {};

    try {
      // Step 1: Process all documents sequentially with detailed progress updates
      const leaseTexts: string[] = [];
      let filesProcessed = 0;
      for (const file of selectedFiles) {
          filesProcessed++;
          const progressPrefix = totalFiles > 1 ? `(${filesProcessed}/${totalFiles})` : '';
          try {
              const text = await extractTextFromPdf(file, (message) => {
                   setProgressMessage(`Processing ${progressPrefix}: ${message}`);
              });
              leaseTexts.push(text);
          } catch (pdfError: any) {
              console.error(`Error extracting text from ${file.name}:`, pdfError);
              throw new Error(`Failed to process "${file.name}". The file may be corrupted, password-protected, or the OCR scan failed.`);
          }
      }

      setProgressMessage('All documents processed. Generating abstract...');
      const leaseText = leaseTexts.join('\n\n--- END OF DOCUMENT ---\n\n');
      leaseTextCache.current = leaseText; // Cache the extracted text

      // Step 2: Call Gemini API for each section
      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY! });
      const baseSystemInstruction = `You are a world-class legal AI specializing in commercial lease abstraction. Your primary objective is to perform a meticulous and comprehensive review of all provided lease documents, which are concatenated and separated by '--- END OF DOCUMENT ---'. You must extract highly detailed information and populate the provided JSON schema with the utmost accuracy.

**Core Directives (Must be followed without exception):**

1.  **Comprehensive Analysis & Chronological Synthesis:**
    *   The documents are provided in chronological order. You MUST analyze them sequentially to build a complete timeline and understanding of the lease.
    *   **Core Principle:** Your final abstract must be a **complete picture** of the current lease agreement. This is achieved by starting with the original lease and layering on the changes from each subsequent amendment.
    *   **Conflict Resolution:** When a later document (e.g., an amendment) explicitly modifies a specific section or clause from an earlier document, the information from the **latest effective amendment** replaces the corresponding older information.
    *   **Information Retention:** Sections or clauses from the original lease or earlier amendments that are **not** explicitly changed by later documents MUST be retained and included in the final abstract. Do not discard information unless it has been directly superseded. For example, if an amendment only changes rent and term dates, you must still extract the 'Use' clause, 'Guaranty', and other unmodified clauses from the original lease.
    *   **Rent Schedules:** The \`baseRentSchedule\` array must be a comprehensive list of all rent schedules defined across the documents. For each entry, you must use the \`sourceDocument\` field to specify which document it came from (e.g., "Original Lease", "First Amendment"). This provides a full history.

2.  **Extreme Detail & Verbatim Extraction:**
    *   Your goal is comprehensiveness, not brevity. **Do not summarize.**
    *   For all clauses, covenants, restrictions, and significant terms (e.g., Use, Exclusivity, Cotenancy, CAM Exclusions), you MUST extract the **full, verbatim text** from the source document. Short, one-sentence summaries are unacceptable.
    *   Provide all available information for every field. If a detail seems minor, include it. The user requires a complete picture.

3.  **Meticulous Sourcing and Citation:**
    *   For every piece of information extracted, you MUST cite its source, including the document name (e.g., "Original Lease," "Second Amendment"), the section number, and the page number. This is non-negotiable for fields where it is applicable.

4.  **CRITICAL: JSON Formatting and Escaping:**
    *   You are generating a JSON response. The string values you provide MUST be correctly escaped to form valid JSON.
    *   **Double Quotes:** Any double quote character (") within the extracted text MUST be escaped with a backslash (e.g., "text with \\"quotes\\""). This is the most common cause of errors.
    *   **Newlines:** Represent newlines within text as \\n.
    *   **Backslashes:** Escape backslashes themselves (e.g., "C:\\\\folder\\\\file").
    *   Failure to produce perfectly valid, parsable JSON will render the entire output useless. Double-check your escaping.

5.  **Absolute Schema Adherence:**
    *   Strictly follow the provided JSON schema, including all data types and formatting rules (Dates: MM/DD/YYYY; Currency: $XXX,XXX.XX; Square Footage: Numeric with commas).

6.  **"Not Provided" as a Last Resort:**
    *   Only use "Not Provided" after you have exhaustively searched all documents and are certain the information does not exist. Before concluding information is missing, double-check all amendments and exhibits.`;

      const sections = Object.entries(leaseAbstractSchema.properties);
      for (const [sectionKey, sectionSchemaDef] of sections) {
        
        if (sectionKey === 'leaseNotes') {
          // Special, granular handling for the complex leaseNotes section
          const leaseNotesData: { [key: string]: any } = {};
          const noteEntries = Object.entries((sectionSchemaDef as any).properties);
  
          for (const [noteKey, noteSchemaDef] of noteEntries) {
            setProgressMessage(`Extracting: Lease Notes (${toTitleCase(noteKey)})...`);
  
            const singleNoteSchema = {
              type: Type.OBJECT,
              properties: { [noteKey]: noteSchemaDef },
            };
  
            const systemInstruction = `${baseSystemInstruction}\n\n**Current Task:** Your sole focus for this request is to extract the data ONLY for the \`${noteKey}\` clause within the Lease Notes. Populate only the fields for this single clause.`;
  
            const response = await ai.models.generateContent({
              model: 'gemini-2.5-flash',
              contents: leaseText,
              config: {
                systemInstruction: systemInstruction,
                responseMimeType: 'application/json',
                responseSchema: singleNoteSchema,
              },
            });
            
            // Handle potentially empty responses gracefully
            if (!response.text || response.text.trim() === '') {
              console.warn(`Received empty response for lease note: ${noteKey}.`);
              if (noteKey === 'Other') {
                  leaseNotesData[noteKey] = [];
              } else {
                  leaseNotesData[noteKey] = { leaseSectionAndPage: "Not Found", notes: "Not Found" };
              }
              continue; // Move to the next note
            }
  
            let jsonResponse;
            try {
              jsonResponse = JSON.parse(response.text);
            } catch (jsonError) {
              console.error(`Failed to parse JSON for note: ${noteKey}. Error:`, jsonError);
              console.error("Malformed response text from Gemini:", response.text);
              throw new Error(`The AI model returned an invalid data format for the "${toTitleCase(noteKey)}" note.`);
            }
  
            Object.assign(leaseNotesData, jsonResponse);
          }
          
          (fullAbstractData as any).leaseNotes = leaseNotesData;
          setAbstractData({ ...fullAbstractData });

        } else {
           // Original logic for all other sections
          setProgressMessage(`Extracting: ${toTitleCase(sectionKey)}...`);

          const sectionSchema = {
            type: Type.OBJECT,
            properties: { [sectionKey]: sectionSchemaDef },
          };

          const systemInstruction = `${baseSystemInstruction}\n\n**Current Task:** Your sole focus for this request is to extract the data ONLY for the \`${sectionKey}\` section. Populate only the fields within this section of the schema.`;
          
          const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: leaseText,
            config: {
              systemInstruction: systemInstruction,
              responseMimeType: 'application/json',
              responseSchema: sectionSchema,
            },
          });

          if (!response.text || response.text.trim() === '') {
             throw new Error(`The AI model returned an empty response for the "${toTitleCase(sectionKey)}" section.`);
          }

          let jsonResponse;
          try {
            jsonResponse = JSON.parse(response.text);
          } catch (jsonError) {
            console.error(`Failed to parse JSON for section: ${sectionKey}. Error:`, jsonError);
            console.error("Malformed response text from Gemini:", response.text);
            throw new Error(`The AI model returned an invalid data format for the "${toTitleCase(sectionKey)}" section.`);
          }
          
          Object.assign(fullAbstractData, jsonResponse);
          setAbstractData({ ...fullAbstractData });
        }
      }

    } catch (e: any) {
      console.error('Operation failed:', e);
      // If a custom error message was thrown, use it. Otherwise, create one.
      const errorMessage = e.message || `An unexpected error occurred. Please check the console and try again.`;
      setError(errorMessage);
    } finally {
      setIsLoading(false);
      setProgressMessage('');
    }
  };

  const handleAskQuestion = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!userQuestion.trim() || isAnswering) return;

    setIsAnswering(true);
    setQaError(null);
    const currentQuestion = userQuestion;
    setChatHistory(prev => [...prev, { role: 'user', content: currentQuestion }]);
    setUserQuestion('');

    try {
      if (!leaseTextCache.current) {
        setProgressMessage('Analyzing documents for Q&A...');
        const leaseTexts = [];
        for (const file of selectedFiles) {
            const text = await extractTextFromPdf(file, (msg) => console.log(msg)); // Use dummy progress reporter
            leaseTexts.push(text);
        }
        leaseTextCache.current = leaseTexts.join('\n\n--- END OF DOCUMENT ---\n\n');
        setProgressMessage('');
      }

      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY! });
      const prompt = `Based ONLY on the content of the following lease document(s), answer the user's question. Be thorough and quote relevant sections if possible. If the answer cannot be found in the documents, state that clearly.\n\n---LEASE DOCUMENTS---\n${leaseTextCache.current}\n\n---USER QUESTION---\n${currentQuestion}`;
      
      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: prompt,
      });

      setChatHistory(prev => [...prev, { role: 'model', content: response.text }]);

    } catch (e: any) {
      console.error('Q&A Gemini call failed:', e);
      setQaError('Sorry, I couldn\'t answer that question. Please try again.');
      // Remove the user's question from history on failure
      setChatHistory(prev => prev.filter(msg => msg.role !== 'user' || msg.content !== currentQuestion));
    } finally {
      setIsAnswering(false);
    }
  };


  const handleExportToCsv = () => {
    if (!abstractData) return;
  
    const rows: (string | number | null | undefined)[][] = [];
  
    const processNode = (data: any, prefix: string) => {
      if (typeof data !== 'object' || data === null) {
        // Handle simple value case
        if (prefix) rows.push([prefix, data]);
        return;
      }
  
      if (Array.isArray(data)) {
        // Handle arrays of objects as a distinct table
        if (data.length > 0 && typeof data[0] === 'object' && data[0] !== null) {
          rows.push([]); // Spacer before table
          if (prefix) rows.push([prefix]); // Sub-section title for the table
          
          // Get all unique keys from all objects in the array to handle inconsistencies
          const allKeys = [...new Set(data.flatMap(item => Object.keys(item)))];
          const headers = allKeys.map(k => toTitleCase(k));
          rows.push(headers);
  
          // Add a row for each object
          data.forEach(item => {
            const row = allKeys.map(k => {
              const val = item[k];
              if (typeof val === 'object' && val !== null) return JSON.stringify(val);
              return val;
            });
            rows.push(row);
          });
        } else {
          // Handle simple arrays or empty arrays
          rows.push([prefix, data.join('; ')]);
        }
      } else { // It's an object
        Object.entries(data).forEach(([key, value]) => {
          const currentKey = prefix ? `${prefix} - ${toTitleCase(key)}` : toTitleCase(key);
          processNode(value, currentKey);
        });
      }
    };
  
    // Iterate over each main section in the abstract data
    Object.entries(abstractData).forEach(([sectionKey, sectionData]) => {
      const sectionTitle = toTitleCase(sectionKey);
      rows.push([sectionTitle]); // Add a title row for the section
      processNode(sectionData, ''); // Process the data within that section
      rows.push([]); // Add a blank line between main sections
    });
  
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const csvOutput = XLSX.utils.sheet_to_csv(ws);
  
    const blob = new Blob([csvOutput], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    if (link.download !== undefined) {
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', 'Lease_Abstract.csv');
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  // Drag and drop handlers
  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); setIsDragging(true); };
  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); setIsDragging(false); };
  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFileChange(e.dataTransfer.files);
      e.dataTransfer.clearData();
    }
  };

  // FIX: Change parameter type to 'any' to handle various data types from the API response and prevent type errors.
  const renderValueHighlight = (value: any) => {
    // FIX: Safely convert value to string, handling null/undefined.
    const stringValue = String(value ?? '');
    if (stringValue.toLowerCase().includes('not provided') || stringValue.toLowerCase().includes('conflicting info') || stringValue.toLowerCase().includes('not found')) {
      return <span className="highlight">{stringValue}</span>;
    }
    return stringValue;
  };
  
  // FIX: Change return type from JSX.Element to React.ReactElement to avoid JSX namespace issues.
  const renderAbstractData = (data: any): React.ReactElement => {
    if (typeof data !== 'object' || data === null) {
      return <div className="value-wrapper">{renderValueHighlight(data)}</div>;
    }
  
    if (Array.isArray(data)) {
      return (
        <div className="result-array">
          {data.map((item, index) => (
            <div key={index} className="result-card nested-card">
              <h4>Item {index + 1}</h4>
              {renderAbstractData(item)}
            </div>
          ))}
        </div>
      );
    }
  
    return (
      <div className="kv-container">
        {Object.entries(data).map(([key, value]) => (
            <div className="kv-row" key={key}>
                <div className="kv-key">{toTitleCase(key)}</div>
                <div className="kv-value">{renderAbstractData(value)}</div>
            </div>
        ))}
      </div>
    );
  };
  

  return (
    <div className="app-container">
      <header>
        <h1>Lease Abstractor</h1>
        <p>AI-Powered Lease Data Extraction</p>
      </header>
      <main className="main-content">
        <div className="input-panel">
          <h2>Lease Documents</h2>
          <div 
            className={`file-dropzone ${isDragging ? 'drag-over' : ''}`}
            onClick={() => fileInputRef.current?.click()}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
            <input
              type="file"
              ref={fileInputRef}
              onChange={(e) => handleFileChange(e.target.files)}
              accept="application/pdf"
              style={{ display: 'none' }}
              disabled={isLoading}
              multiple
            />
            {selectedFiles.length > 0 ? (
              <div className="file-list-container">
                <ul className="file-list">
                  {selectedFiles.map((file, index) => (
                    <li key={file.name} className="file-item">
                      <i className="fa-solid fa-file-pdf"></i>
                      <span>{file.name}</span>
                      <button 
                        className="remove-file-btn"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleRemoveFile(index);
                        }}
                        aria-label={`Remove ${file.name}`}
                      >
                        &times;
                      </button>
                    </li>
                  ))}
                </ul>
                <button
                  className="clear-all-btn"
                  onClick={(e) => {
                    e.stopPropagation();
                    setSelectedFiles([]);
                    setAbstractData(null);
                    setChatHistory([]);
                    leaseTextCache.current = '';
                    if (fileInputRef.current) fileInputRef.current.value = '';
                  }}
                  aria-label="Remove all files"
                >
                  Clear All
                </button>
              </div>
            ) : (
              <div className="drop-prompt">
                 <i className="fa-solid fa-cloud-arrow-up"></i>
                 <p>Drag & drop your PDFs here, or <strong>click to browse</strong>.</p>
              </div>
            )}
          </div>
          <button onClick={handleGenerateAbstract} disabled={selectedFiles.length === 0 || isLoading} aria-live="polite">
            {isLoading ? (
              <>
                <div className="spinner" aria-hidden="true"></div>
                {progressMessage || 'Generating...'}
              </>
            ) : (
              'Generate Abstract'
            )}
          </button>
        </div>
        <div className="output-panel">
          <div className="output-header">
            <h2>Lease Abstract</h2>
            {abstractData && !isLoading && (
              <button className="export-button" onClick={handleExportToCsv}>
                <i className="fa-solid fa-file-csv"></i> Export to CSV
              </button>
            )}
          </div>
          <div className="results-container" aria-live="polite">
            {isLoading && !abstractData && (
               <div className="skeleton-loader">
                 <div className="skeleton-card"></div>
                 <div className="skeleton-card"></div>
                 <div className="skeleton-card"></div>
               </div>
            )}
            {error && <div className="error-message" role="alert">{error}</div>}
            {abstractData ? (
              Object.entries(abstractData).map(([sectionTitle, sectionData]) => (
                <div key={sectionTitle} className="result-card">
                  <h3>
                    <i className={sectionIcons[sectionTitle] || 'fa-solid fa-file-contract'}></i>
                    {toTitleCase(sectionTitle)}
                  </h3>
                  {renderAbstractData(sectionData)}
                </div>
              ))
            ) : (
              !isLoading && !error && <p className="placeholder-text">Your extracted lease details will appear here.</p>
            )}
          </div>

          {selectedFiles.length > 0 && !isLoading && (
            <div className="qa-panel">
              <div className="output-header">
                <h2>Ask a Question</h2>
              </div>
              <div className="chat-history" ref={chatHistoryRef}>
                {chatHistory.length === 0 && <p className="placeholder-text">Ask a question to get started.</p>}
                {chatHistory.map((msg, index) => (
                  <div key={index} className={`chat-message ${msg.role}-message`}>
                    <p>{msg.content}</p>
                  </div>
                ))}
                {isAnswering && (
                  <div className="chat-message model-message loading-message">
                    <div className="spinner"></div>
                  </div>
                )}
                {qaError && <div className="error-message" role="alert">{qaError}</div>}
              </div>
              <form className="qa-input-form" onSubmit={handleAskQuestion}>
                <textarea
                  value={userQuestion}
                  onChange={(e) => setUserQuestion(e.target.value)}
                  onKeyDown={(e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleAskQuestion(e); }}}
                  placeholder="e.g., Pull up all references to CAM"
                  disabled={isAnswering || isLoading}
                  rows={1}
                  aria-label="Ask a question about the documents"
                />
                <button type="submit" disabled={isAnswering || isLoading || !userQuestion.trim()} aria-label="Send question">
                  <i className="fa-solid fa-paper-plane"></i>
                </button>
              </form>
            </div>
          )}
        </div>
      </main>
    </div>
  );
};

const root = ReactDOM.createRoot(document.getElementById('root')!);
root.render(<App />);