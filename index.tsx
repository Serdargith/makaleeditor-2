
import React, { useState, useRef, useEffect } from "react";
import { createRoot } from "react-dom/client";
import { GoogleGenAI } from "@google/genai";
import Markdown from 'react-markdown';
import * as XLSX from 'xlsx';

// Define styles
const styles = {
  container: {
    fontFamily: "'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
    maxWidth: "1200px",
    margin: "0 auto",
    padding: "20px",
    backgroundColor: "#f9f9f9",
    minHeight: "100vh",
    color: "#333",
  },
  header: {
    backgroundColor: "#005f73",
    color: "white",
    padding: "20px",
    borderRadius: "8px",
    marginBottom: "20px",
    boxShadow: "0 4px 6px rgba(0,0,0,0.1)",
  },
  title: {
    margin: 0,
    fontSize: "24px",
    fontWeight: "600",
  },
  subtitle: {
    margin: "5px 0 0 0",
    fontSize: "14px",
    opacity: 0.9,
  },
  mainContent: {
    display: "flex",
    gap: "20px",
    flexDirection: "column" as const,
  },
  inputSection: {
    backgroundColor: "white",
    padding: "20px",
    borderRadius: "8px",
    boxShadow: "0 2px 4px rgba(0,0,0,0.05)",
    display: "flex",
    flexDirection: "column" as const,
    gap: "15px",
  },
  fileUploadSection: {
    display: "flex",
    gap: "20px",
    flexWrap: "wrap" as const,
    marginBottom: "10px",
  },
  fileInputGroup: {
    flex: "1 1 300px",
    display: "flex",
    flexDirection: "column" as const,
    gap: "8px",
  },
  fileList: {
    fontSize: "12px",
    color: "#555",
    marginTop: "5px",
    paddingLeft: "10px",
    borderLeft: "2px solid #0a9396",
  },
  textarea: {
    width: "100%",
    minHeight: "300px",
    padding: "15px",
    border: "1px solid #ddd",
    borderRadius: "4px",
    fontSize: "14px",
    fontFamily: "monospace",
    resize: "vertical" as const,
  },
  button: {
    backgroundColor: "#0a9396",
    color: "white",
    border: "none",
    padding: "12px 24px",
    fontSize: "16px",
    fontWeight: "bold",
    borderRadius: "4px",
    cursor: "pointer",
    transition: "background-color 0.2s",
    alignSelf: "flex-start",
  },
  buttonDisabled: {
    backgroundColor: "#94d2bd",
    cursor: "not-allowed",
  },
  outputSection: {
    backgroundColor: "white",
    padding: "40px",
    borderRadius: "8px",
    boxShadow: "0 2px 4px rgba(0,0,0,0.05)",
    lineHeight: "1.6",
  },
  loading: {
    textAlign: "center" as const,
    padding: "40px",
    color: "#005f73",
    fontSize: "18px",
    fontStyle: "italic",
  },
  error: {
    backgroundColor: "#ffe8e8",
    color: "#c0392b",
    padding: "15px",
    borderRadius: "4px",
    marginBottom: "20px",
  },
  label: {
    fontWeight: "bold",
    marginBottom: "5px",
    display: "block",
  },
  helperText: {
    fontSize: "12px",
    color: "#666",
    marginBottom: "5px",
  },
  fileInput: {
    padding: "8px",
    border: "1px dashed #ccc",
    borderRadius: "4px",
    backgroundColor: "#fafafa",
  }
};

const EDITOR_PROMPT = `
Act as a world-class Baş Editör specializing in Ortopedi ve Travmatoloji. Given the following context, criteria, and instructions, perform a thorough, literature-informed editorial review of the provided academic manuscript and deliver detailed, actionable feedback suitable for communicating with authors and making an editorial decision.

## Context
- Hedef: Dünya çapında saygın bir SCIE dergisinin baş editörü rolünde, sunulan ortopedi ve travmatoloji makalesini güncel literatür verileriyle karşılaştırarak değerlendirmek ve ayrıntılı geri bildirim sağlamak.
- Hedef okuyucu: akademik hakemler, dergi editörleri ve makale yazarları.
- Kaynak beklentisi: En güncel literatür (tercihen son 5 yıl; en az 10 uygun, yüksek etkili referans), uygun rehberler (CONSORT, STROBE, PRISMA, CARE vb.) ve metodolojik standartlarla karşılaştırma.

## Approach
- İlk adım olarak makalenin başlığı, özet, amaç/hypotez, çalışma tasarımı, hasta popülasyonu, müdahaleler, sonuçlar, istatistiksel analizler, sonuçlar ve tartışma bölümlerini kısa ve net bir şekilde özetle.
- Çalışma tipini otomatik olarak belirle ve ilgili raporlama rehberi (CONSORT, STROBE, PRISMA, CARE vs.) üzerinden uyumluluğu kontrol et.
- Metodoloji, hasta seçimi, dahil/hariç etme kriterleri, rasgeleleştirme/blinding (varsa), takip süresi, güç analizi ve örnekleme boyutu açısından sağlamlık analizi yap.
- İstatistiksel analizleri satır satır değerlendir; kullanılan testlerin uygunluğunu, varsayımların kontrolünü, çoklu karşılaştırma düzeltmelerini, etki büyüklüklerinin ve güven aralıklarının raporlanmasını denetle. Gerekirse alternatif analizler ve ek tablolar/analizler öner.
- Sonuç yorumlarını literatürle karşılaştır ve hangi bulguların yeni/önemli olduğunu, hangilerinin literatürde daha önce gösterildiğini belirle.
- Çalışmanın klinik önemini ve uygulamaya etkisini değerlendir; potansiyel çarpıtmalar, önyargılar, sınırlamalar ve bunların nasıl giderilebileceğine dair somut öneriler sun.
- Etik onay, hasta onamı, çatışma beyanları ve finansman açıklamalarını kontrol et.
- Makalenin dil, yapı, tablo/şekil kalitesi ve referans formatı açısından düzeltme ve yeniden yazım önerileri hazırla.
- Güncel literatür taraması yap (PubMed/Scopus/Embase tercihli) ve en az 10 son döneme ait, uygun ve yüksek kaliteli referansı tam bibliyografik biçimde (Vancouver stili, DOI ile) ekle. Bu referanslar makaledeki iddiaları desteklemeli veya tartışmayı zenginleştirmeli.
- Yapılacak değişiklikleri önem sırasına göre sınıflandır (Acil/Öncelikli — Major; İyileştirme — Minor) ve her biri için beklenen yazar yanıtı formatı belirt.
- Editöre önerilecek nihai karar önerisini net bir gerekçe ile sun: Kabul, Küçük revizyon, Büyük revizyon, Reddetme. Her bir karar için gereken en az değişiklikleri açıkça belirt.

## Response Format
- 1) Kısa Özet: 3–6 cümlelik net özet (araştırma sorusu, yöntem, ana bulgu, sonuç).
- 2) Genel Değerlendirme ve Yenilik/Önem (2–4 kısa paragraf).
- 3) Ana Endişeler (Major Concerns) — numaralandırılmış, her biri: Problem açıklaması; neden önemli olduğu; yazarların ne yapması gerektiği; örnek düzeltme/metin veya ek analiz önerisi.
- 4) Küçük Endişeler (Minor Concerns) — numaralandırılmış, kısa düzeltme önerileri (dil, biçim, referans eksiklikleri, küçük metodolojik notlar).
- 5) Metodoloji ve İstatistik Detayları — kullanılan testlere dair teknik değerlendirme; gerekli alternatif analizler veya ek tablolar; örnek R/Python kodu (kısaca) veya istatistik komutları gerekiyorsa (opsiyonel).
- 6) Literatür Bağlamı ve Eksik/Kritik Referanslar — en az 10 güncel referans (Vancouver stili, DOI ile), her birinin kısa bir açıklaması (2–3 cümle) ve makaledeki hangi iddia/alanı desteklediği.
- 7) Şekiller/Tablolar — her bir figure/table için açıklanan sorunlar ve yeniden düzenleme önerileri.
- 8) Etik, Uyumluluk ve Bildirimler — etik onay durumu, onam, veri paylaşımı beyanı, çıkar çatışması kontrolü.
- 9) Önerilen Başlık/Özet/Keywords — bir veya iki alternatif başlık; yeniden yazılmış 200–250 kelimelik yapılandırılmış özet (Introduction, Methods, Results, Conclusion) ve 6–8 anahtar kelime.
- 10) Örnek Hakem Raporu ve Yazarlar İçin Kısa Revizyon Notu — editöre sunulacak kısa karar cümlesi ve yazarların yanıtlaması gereken maddeleri içeren net revizyon listesi.
- 11) Nihai Editoryal Karar Önerisi — karar tercihi (Kabul/Küçük revizyon/Büyük revizyon/Reddetme) ve kısa gerekçe.
- 12) Tahmini Revizyon Süresi ve Önceliklendirme.
- 13) Referans Listesi — metinde atıfta bulunulan tüm kaynakların tam bibliyografik listesi (Vancouver stili, DOI).

Tüm bölümler açık başlıklarla ayrılmalı ve numaralandırılmalı. Dil: Tüm geri bildirim Türkçe yazılmalı.
`;

const App = () => {
  const [manuscript, setManuscript] = useState("");
  const [figures, setFigures] = useState<File[]>([]);
  const [tables, setTables] = useState<File[]>([]);
  const [dataFiles, setDataFiles] = useState<File[]>([]);
  
  const [result, setResult] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>, setFiles: React.Dispatch<React.SetStateAction<File[]>>) => {
    if (e.target.files) {
      setFiles(Array.from(e.target.files));
    }
  };

  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        // Remove data URL prefix (e.g. "data:image/png;base64,")
        const base64 = result.split(',')[1];
        resolve(base64);
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  };

  const parseExcelToCSV = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const csv = XLSX.utils.sheet_to_csv(worksheet);
          resolve(csv);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  };

  const handleAnalyze = async () => {
    if (!manuscript.trim() && figures.length === 0 && tables.length === 0 && dataFiles.length === 0) {
      setError("Please provide manuscript text or attach files.");
      return;
    }

    setLoading(true);
    setError("");
    setResult("");

    try {
      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
      const parts: any[] = [];

      // 1. Add Prompt + Manuscript Text
      const fullPrompt = `${EDITOR_PROMPT}\n\n--- MANUSCRIPT START ---\n${manuscript}\n--- MANUSCRIPT END ---`;
      parts.push({ text: fullPrompt });

      // 2. Process Figures (Images)
      for (const file of figures) {
        try {
          const base64Data = await fileToBase64(file);
          parts.push({
            inlineData: {
              mimeType: file.type,
              data: base64Data
            }
          });
          parts.push({ text: `[Attached Figure: ${file.name}]` });
        } catch (e) {
          console.error(`Failed to process figure ${file.name}`, e);
        }
      }

      // 3. Process Tables (Images or PDF)
      for (const file of tables) {
        try {
          const base64Data = await fileToBase64(file);
          parts.push({
            inlineData: {
              mimeType: file.type,
              data: base64Data
            }
          });
          parts.push({ text: `[Attached Table: ${file.name}]` });
        } catch (e) {
          console.error(`Failed to process table ${file.name}`, e);
        }
      }

      // 4. Process Data Files (Excel/CSV)
      for (const file of dataFiles) {
        try {
          let csvContent = "";
          if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            csvContent = await parseExcelToCSV(file);
          } else {
            // Assume CSV or text
            csvContent = await file.text();
          }
          // Truncate if too long to avoid huge context usage, though Gemini Pro has large context.
          // Let's limit reasonably just in case, e.g. 50k chars
          const truncatedContent = csvContent.slice(0, 100000); 
          parts.push({ text: `\n--- DATA FILE CONTENT (${file.name}) ---\n${truncatedContent}\n--- END DATA FILE ---\n` });
        } catch (e) {
           console.error(`Failed to process data file ${file.name}`, e);
           parts.push({ text: `[Error reading data file ${file.name}]` });
        }
      }

      const response = await ai.models.generateContent({
        model: 'gemini-3-pro-preview',
        contents: { parts: parts },
        config: {
          tools: [{ googleSearch: {} }],
          thinkingConfig: { thinkingBudget: 2048 },
        }
      });

      const text = response.text;
      setResult(text);
      
    } catch (err: any) {
      console.error(err);
      setError("An error occurred during analysis: " + (err.message || "Unknown error"));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={styles.container}>
      <header style={styles.header}>
        <h1 style={styles.title}>Orthopedic Journal Chief Editor AI</h1>
        <p style={styles.subtitle}>
          Automated Academic Editorial Review for Orthopedics & Traumatology
        </p>
      </header>

      <main style={styles.mainContent}>
        <div style={styles.inputSection}>
          {error && <div style={styles.error}>{error}</div>}
          
          <label style={styles.label}>1. Manuscript Text</label>
          <div style={styles.helperText}>
            Paste the full text of the article (Abstract, Introduction, Methods, Results, Discussion) here.
          </div>
          <textarea
            style={styles.textarea}
            value={manuscript}
            onChange={(e) => setManuscript(e.target.value)}
            placeholder="Paste title, abstract, and body text here..."
          />

          <label style={{...styles.label, marginTop: '20px'}}>2. Attachments</label>
          <div style={styles.fileUploadSection}>
            <div style={styles.fileInputGroup}>
              <label style={styles.helperText}>Figures (Images: .png, .jpg, .webp)</label>
              <input 
                type="file" 
                multiple 
                accept="image/png, image/jpeg, image/webp"
                onChange={(e) => handleFileChange(e, setFigures)}
                style={styles.fileInput}
              />
              <div style={styles.fileList}>
                {figures.map((f, i) => <div key={i}>{f.name}</div>)}
              </div>
            </div>

            <div style={styles.fileInputGroup}>
              <label style={styles.helperText}>Tables (Images or PDF)</label>
              <input 
                type="file" 
                multiple 
                accept="image/*, application/pdf"
                onChange={(e) => handleFileChange(e, setTables)}
                style={styles.fileInput}
              />
              <div style={styles.fileList}>
                {tables.map((f, i) => <div key={i}>{f.name}</div>)}
              </div>
            </div>

            <div style={styles.fileInputGroup}>
              <label style={styles.helperText}>Supplementary Data (Excel .xlsx, .csv)</label>
              <input 
                type="file" 
                multiple 
                accept=".xlsx, .xls, .csv, .txt"
                onChange={(e) => handleFileChange(e, setDataFiles)}
                style={styles.fileInput}
              />
              <div style={styles.fileList}>
                {dataFiles.map((f, i) => <div key={i}>{f.name}</div>)}
              </div>
            </div>
          </div>

          <button
            style={{
              ...styles.button,
              ...(loading ? styles.buttonDisabled : {}),
            }}
            onClick={handleAnalyze}
            disabled={loading}
          >
            {loading ? "Analyzing Manuscript & Data..." : "Perform Editorial Review"}
          </button>
        </div>

        {loading && (
          <div style={styles.loading}>
            <p>Processing manuscript and attachments...</p>
            <p>Performing literature search and methodological review...</p>
          </div>
        )}

        {result && (
          <div style={styles.outputSection}>
            <Markdown>{result}</Markdown>
          </div>
        )}
      </main>
    </div>
  );
};

const root = createRoot(document.getElementById("root"));
root.render(<App />);
