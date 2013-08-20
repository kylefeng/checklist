(ns checklist.core
  (:require [clojure.java.io :as io]
            [clojure.string :as str])
  (:use [clojure.pprint]
        [clojure.tools.cli :only [cli]]
        [org.satta.glob]
        [dk.ative.docjure.spreadsheet])
  (:import [org.apache.poi.ss.usermodel CellStyle])
  (:gen-class true))

(defn- load-md-files
  "Load all markdown files in the specified directory,
   and return a lazy seq contains every line, file per vector."
  [dir]
  (for [md (glob (str dir "/*.md"))]
    (with-open [rdr (io/reader md)]
       (reduce
       #(if (not (str/blank? %2))
          (conj %1 (str/trim %2)) %1) [] (line-seq rdr)))))

(defn- make-checklist [title]
  (atom {:title title :items (sorted-map)}))

(defn- add-item!
  ([ckl name] (swap! ckl assoc-in [:items name] []) ckl)
  ([ckl name item](swap! ckl update-in [:items name] conj item) ckl))

(defn- parse-md-vec
  "Parse a markdown lines vector into the ArrayList of domain models."
   [md-vec]
   (let [cl-stack (fn [stack]
                    (if (> (.size stack) 0)
                      (.pop stack)))
         extract (fn [re line]
                   (-> line
                       (str/split re)
                       rest
                       first
                       str/trim))
         stack-of-ckl (java.util.Stack.)
         stack-of-item (java.util.Stack.)
         result (java.util.ArrayList.)]
     (doseq [line md-vec]
       (cond
        (re-find #"^\* .+" line) (let [line (extract #"\*" line)]
                                   (cl-stack stack-of-item)
                                   (add-item! (.peek stack-of-ckl) line)
                                   (.push stack-of-item line))
        (re-find #"^- .+" line) (let [line (extract #"-" line)]
                                  (add-item! (.peek stack-of-ckl)
                                             (.peek stack-of-item)
                                             line)) 
        :else (when-not (re-find #"^-+" line)
                (do
                  (cl-stack stack-of-ckl)
                  (let [ckl (make-checklist line)]
                    (.push stack-of-ckl ckl)
                    (.add result ckl))))))
     result))

(defn- gen-excel-rows
  "Generate excel rows from parsed markdow files data."
  [^java.util.List parsed]
  (let [rows (atom [])]
    (doseq [ckl parsed]
      (let [title (:title @ckl)
            items (:items @ckl)
            ks (keys items)]
        (doseq [k ks]
          (let [v (get items k)]
            (if (> (count v) 0)
              (doseq [item v]
                (swap! rows conj [title k item])))))))
    @rows))

(defn- gen-excel [file-name rows]
  (let [rows (cons ["大类" "小类" "检查点"] rows)
        sheet-name "checklist"
        wb (create-workbook sheet-name rows)
        sheet (select-sheet sheet-name wb)
        header-row (first (row-seq sheet))
        style (.createCellStyle wb)
        len-of-char (fn [n] (* n 256))]
    (do
      (doto sheet
        (.setColumnWidth 0 (len-of-char 25))
        (.setColumnWidth 1 (len-of-char 35))
        (.setColumnWidth 2 (len-of-char 60))
        (.createFreezePane 3 1))
      (doto style
        (.setWrapText true)
        (.setVerticalAlignment CellStyle/VERTICAL_CENTER)
        (.setBorderBottom      CellStyle/BORDER_THIN)
        (.setBorderLeft        CellStyle/BORDER_THIN)
        (.setBorderTop         CellStyle/BORDER_THIN)
        (.setBorderRight       CellStyle/BORDER_THIN))
      (dorun (map #(set-row-style! % style) (rest (row-seq sheet))))
      (set-row-style! header-row (create-cell-style! wb {:background :yellow,
                                                         :font {:bold true}}))
      (save-workbook! file-name wb))))

(defn checklist-to-excel [dir file-name]
  (->> (load-md-files dir)
       (map parse-md-vec)
       (map gen-excel-rows)
       (reduce into [])
       (gen-excel file-name)))

(defn -main [& args]
  (let [[opts args banner]
        (cli args)
        dir (first args)
        filename (second args)]
    (if-not (and
             (= (count args) 2)
             (.endsWith filename ".xlsx")) 
      (do
        (println "")
        (println "Usage: checklist <markdown files dir> <filename for generated excel (.xlsx)>"))
      (checklist-to-excel dir filename))))
