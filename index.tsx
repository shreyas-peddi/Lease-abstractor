/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import ReactDOM from 'react-dom/client';
import { GoogleGenAI, Type } from '@google/genai';
import * as pdfjsLib from 'pdfjs-dist';
import * as XLSX from 'xlsx';

// Set up the PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://aistudiocdn.com/pdfjs-dist@^4.5.136/build/pdf.worker.min.mjs`;

// Helper to convert camelCase to Title Case for display
const toTitleCase = (str: string) => {
  if (!str) return '';
  return str.replace(/([A-Z])/g, ' $1').replace(/^./, (s) => s.toUpperCase());
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
          issues: { type: Type.STRING },
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
        leaseSectionAndPage: { type: Type.STRING },
        baseRentSchedule: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              incomeCategory: { type: Type.STRING },
              effectiveDate: { type: Type.STRING },
              endDate: { type: Type.STRING },
              annualAmountPerSf: { type: Type.STRING },
              annualTotal: { type: Type.STRING },
              monthlyAmount: { type: Type.STRING },
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

    keyClauses: {
      type: Type.OBJECT,
      properties: {
        cotenancyRequirements: clauseSchema,
        landlordKickout: clauseSchema,
        tenantKickout: clauseSchema,
        tenantGoDark: clauseSchema,
        landlordRestrictions: clauseSchema,
        assignmentAndSubletting: clauseSchema,
        shoppingCenterAlterations: clauseSchema,
        operatingCovenant: clauseSchema,
        lateChargesNSFFee: clauseSchema,
        defaultClause: clauseSchema,
        guaranty: clauseSchema,
        purchaseOptionROFR: clauseSchema,
        marketingOrPromotionalFee: clauseSchema,
        holdoverTerms: clauseSchema,
        signage: clauseSchema,
        estoppel: clauseSchema,
        eminentDomainAndSubordination: clauseSchema,
        damageOrDestruction: clauseSchema,
        relocationRight: clauseSchema,
        additionalNotes: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              leaseSectionAndPage: { type: Type.STRING },
              reference: { type: Type.STRING },
              notes: { type: Type.STRING },
            },
          },
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
    }
  };
  
  const extractTextFromPdf = async (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (event) => {
        if (!event.target?.result) {
          return reject(new Error("Failed to read file."));
        }
        try {
          const typedArray = new Uint8Array(event.target.result as ArrayBuffer);
          const pdf = await pdfjsLib.getDocument(typedArray).promise;
          let fullText = '';
          for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map(item => ('str' in item ? item.str : '')).join(' ');
            fullText += pageText + '\n\n';
          }
          resolve(fullText);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    });
  };

  const handleGenerateAbstract = async () => {
    if (selectedFiles.length === 0) {
      setError('Please upload one or more lease documents.');
      return;
    }
    setIsLoading(true);
    setError(null);
    setAbstractData(null);
    setProgressMessage('Preparing documents...');

    const fullAbstractData = {};

    try {
      const leaseTexts = await Promise.all(
        selectedFiles.map(file => extractTextFromPdf(file))
      );
      const leaseText = leaseTexts.join('\n\n--- END OF DOCUMENT ---\n\n');

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

4.  **Absolute Schema Adherence:**
    *   Strictly follow the provided JSON schema, including all data types and formatting rules (Dates: MM/DD/YYYY; Currency: $XXX,XXX.XX; Square Footage: Numeric with commas).

5.  **"Not Provided" as a Last Resort:**
    *   Only use "Not Provided" after you have exhaustively searched all documents and are certain the information does not exist. Before concluding information is missing, double-check all amendments and exhibits.`;

      const sections = Object.entries(leaseAbstractSchema.properties);
      for (const [sectionKey, sectionSchemaDef] of sections) {
        setProgressMessage(`Extracting: ${toTitleCase(sectionKey)}...`);

        const sectionSchema = {
          type: Type.OBJECT,
          properties: {
            [sectionKey]: sectionSchemaDef,
          },
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

        const jsonResponse = JSON.parse(response.text);
        Object.assign(fullAbstractData, jsonResponse);
        setAbstractData({ ...fullAbstractData });
      }

    } catch (e: any) {
      console.error('Gemini API call failed:', e);
      const currentSection = progressMessage.replace('Extracting: ', '').replace('...', '');
      setError(`Failed to extract data for "${currentSection}". Please try again.`);
    } finally {
      setIsLoading(false);
      setProgressMessage('');
    }
  };

  const handleExportToExcel = () => {
    if (!abstractData) return;

    const flattenDataForExcel = (data: any) => {
      const rows: (string | number)[][] = [];
      const process = (obj: any, prefix = '') => {
        if (typeof obj !== 'object' || obj === null) return;

        Object.entries(obj).forEach(([key, value]) => {
          const newPrefix = prefix ? `${prefix} > ${toTitleCase(key)}` : toTitleCase(key);
          if (typeof value === 'object' && value !== null) {
            if (Array.isArray(value)) {
              rows.push([newPrefix]);
              value.forEach((item, index) => {
                rows.push([`${newPrefix} [${index + 1}]`]);
                process(item, `  `); // Indent array items
              });
            } else {
              process(value, newPrefix);
            }
          } else {
            // FIX: Explicitly convert the value to a string to prevent type errors when pushing to the `rows` array, which expects `string | number`.
            rows.push([newPrefix, String(value ?? '')]);
          }
        });
      };
      process(data);
      return rows;
    };

    const data = flattenDataForExcel(abstractData);
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 45 }, { wch: 60 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Lease Abstract');
    XLSX.writeFile(wb, 'Lease_Abstract.xlsx');
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
    if (stringValue.toLowerCase().includes('not provided') || stringValue.toLowerCase().includes('conflicting info')) {
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
      <ul className="result-list">
        {Object.entries(data).map(([key, value]) => (
          <li key={key}>
            <strong>{toTitleCase(key)}:</strong>
            {renderAbstractData(value)}
          </li>
        ))}
      </ul>
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
              <button className="export-button" onClick={handleExportToExcel}>
                <i className="fa-solid fa-file-excel"></i> Export to Excel
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
                  <h3>{toTitleCase(sectionTitle)}</h3>
                  {renderAbstractData(sectionData)}
                </div>
              ))
            ) : (
              !isLoading && !error && <p className="placeholder-text">Your extracted lease details will appear here.</p>
            )}
          </div>
        </div>
      </main>
    </div>
  );
};

const root = ReactDOM.createRoot(document.getElementById('root')!);
root.render(<App />);