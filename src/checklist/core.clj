(ns checklist.core
  (:require [clojure.java.io :as io]
            [clojure.string :as str])
  (:use [clojure.pprint]
        [org.satta.glob]
        [dk.ative.docjure.spreadsheet]))

(defn- load-md-files
  "Load all markdown files in the specified directory,
   and return a lazy seq contains every line, file per vector."
  [dir]
  (for [md (glob (str dir "/*.md"))]
    (with-open [rdr (io/reader md)]
       (reduce
       #(if (not (str/blank? %2))
          (conj %1 (str/trim %2)) %1) [] (line-seq rdr)))))

(defn- make-check-list [title]
  (atom {:title title :items {}}))

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
                  (let [ckl (make-check-list line)]
                    (.push stack-of-ckl ckl)
                    (.add result ckl))))))
     result))

(defn- gen-excel
  [^java.util.List parsed]
  (doseq [ckl parsed]
    ))
