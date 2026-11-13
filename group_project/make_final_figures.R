# Figures for Group_B_Final_Answers.pptx — same pipeline as the final HTML
suppressPackageStartupMessages({
  library(tidyverse); library(here); library(microViz); library(phyloseq)
  library(vegan); library(patchwork); library(DESeq2); library(ggrepel)
})
source(here("R/count_tables.R"))
set.seed(1)
dir.create(here("group_project/slides_assets"), showWarnings = FALSE)

d  <- read_counts("group_b")
ps <- as_counts("group_b", table = "microbiome", min_count = 10, min_samples = 20)
stopifnot(identical(colnames(d$species), d$metadata$sample_id))

# ---- Figure 1: microbiome ordination, arm vs BV (Part 1) --------------------
o <- ps |> tax_transform("identity") |> dist_calc("bray") |> ord_calc(method = "PCoA")
eig <- ord_get(o)$CA$eig
lab1 <- sprintf("Axis 1: %.1f%%", eig[1]/sum(eig)*100)
lab2 <- sprintf("Axis 2: %.1f%%", eig[2]/sum(eig)*100)

p_arm <- ord_plot(o, colour = "arm", size = 2.4) +
  stat_ellipse(aes(colour = arm), linetype = 2, linewidth = 0.5) +
  scale_colour_manual(values = c(pill = "#E69F00", injectable = "#56B4E9", ring = "#009E73")) +
  labs(title = "Coloured by contraception arm", colour = NULL,
       x = lab1, y = lab2) +
  theme_bw(base_size = 12) + theme(legend.position = "bottom")

p_bv <- ord_plot(o, colour = "bv", size = 2.4) +
  stat_ellipse(aes(colour = bv), linetype = 2, linewidth = 0.5) +
  scale_colour_manual(values = c(negative = "#999999", intermediate = "#CC79A7", positive = "#D55E00")) +
  labs(title = "Coloured by BV status", colour = NULL,
       x = lab1, y = lab2) +
  theme_bw(base_size = 12) + theme(legend.position = "bottom")

fig1 <- p_arm | p_bv
ggsave(here("group_project/slides_assets/fig1_ordination.png"), fig1,
       width = 9.4, height = 4.9, dpi = 300, bg = "white")

# ---- Part 2 DESeq2: volcano + top-gene boxplots (Figure 2) ------------------
meta_d <- d$metadata |>
  mutate(bv = factor(bv), rnaseq_batch = factor(rnaseq_batch),
         delivery = factor(ifelse(arm == "ring", "local", "systemic"),
                           levels = c("systemic", "local")))
keep <- rowSums(d$genes >= 10) >= 23
dds <- DESeq(DESeqDataSetFromMatrix(d$genes[keep, ], meta_d,
                                    design = ~ rnaseq_batch + bv + delivery), quiet = TRUE)
res <- results(dds, contrast = c("delivery", "local", "systemic"))
sig <- as.data.frame(res) |> rownames_to_column("ensembl_gene_id") |> filter(padj < 0.05)
rd  <- as.data.frame(res) |> rownames_to_column("ensembl_gene_id") |>
  left_join(d$gene_info[, c("ensembl_gene_id", "symbol")], by = "ensembl_gene_id")
top10 <- head(arrange(sig, padj), 10)$ensembl_gene_id

volcano <- ggplot(rd, aes(log2FoldChange, -log10(pvalue))) +
  geom_point(aes(colour = padj < 0.05), size = 0.8, alpha = 0.5) +
  geom_vline(xintercept = 0, linetype = "dashed", colour = "grey50") +
  geom_hline(yintercept = -log10(0.05), linetype = "dashed", colour = "grey50") +
  geom_label_repel(data = filter(rd, ensembl_gene_id %in% top10),
                   aes(label = symbol), size = 3.2, min.segment.length = 0) +
  scale_colour_manual(values = c(`FALSE` = "grey70", `TRUE` = "#D55E00"),
                      labels = c(`FALSE` = "ns", `TRUE` = "FDR < 5%"), name = NULL) +
  labs(x = "log2 fold change (ring vs systemic)", y = "-log10(p-value)") +
  theme_bw(base_size = 12) + theme(legend.position = "bottom")

top6 <- head(arrange(sig, padj), 4)$ensembl_gene_id
norm <- log10(counts(dds, normalized = TRUE)[top6, ] + 1) |>
  as.data.frame() |> rownames_to_column("ensembl_gene_id") |>
  pivot_longer(-ensembl_gene_id, names_to = "sample_id", values_to = "log10_count") |>
  left_join(meta_d[, c("sample_id", "delivery")], by = "sample_id") |>
  left_join(d$gene_info[, c("ensembl_gene_id", "symbol")], by = "ensembl_gene_id")

boxes <- ggplot(norm, aes(delivery, log10_count, fill = delivery)) +
  geom_boxplot(outlier.size = 0.8) +
  facet_wrap(~ symbol, scales = "free_y") +
  scale_fill_manual(values = c(local = "#D55E00", systemic = "#56B4E9")) +
  labs(x = NULL, y = "log10 normalised count") +
  theme_bw(base_size = 12) + theme(legend.position = "none")

fig2 <- volcano | boxes
ggsave(here("group_project/slides_assets/fig2_volcano.png"), fig2,
       width = 10.4, height = 4.6, dpi = 300, bg = "white")

# ---- Part 3: head-to-head barplot (Figure 3) --------------------------------
lacto_species <- d$taxonomy$taxon[d$taxonomy$Genus == "Lactobacillus"]
frac <- colSums(d$species[lacto_species, ]) / colSums(d$species)
meta_c <- d$metadata |>
  mutate(lacto_fraction = frac[sample_id],
         community = factor(ifelse(lacto_fraction > 0.5, "lactobacillus", "anaerobe"),
                            levels = c("lactobacillus", "anaerobe")),
         delivery  = factor(ifelse(arm == "ring", "local", "systemic"),
                            levels = c("systemic", "local")))
dds_c <- DESeq(DESeqDataSetFromMatrix(d$genes[keep, ], meta_c,
                                      design = ~ rnaseq_batch + delivery + community),
               quiet = TRUE)
r_com <- results(dds_c, contrast = c("community", "anaerobe", "lactobacillus"))
r_del <- results(dds_c, contrast = c("delivery", "local", "systemic"))

counts_cmp <- tibble(
  explanation = c("Community type", "Method (ring vs systemic)"),
  `raw p < 0.05` = c(sum(r_com$pvalue < 0.05, na.rm = TRUE),
                     sum(r_del$pvalue < 0.05, na.rm = TRUE)),
  `FDR < 5%` = c(sum(r_com$padj < 0.05, na.rm = TRUE),
                 sum(r_del$padj < 0.05, na.rm = TRUE)))

fig3 <- ggplot(pivot_longer(counts_cmp, -explanation),
               aes(explanation, value, fill = name)) +
  geom_col(position = position_dodge(0.75), width = 0.62) +
  geom_text(aes(label = value), position = position_dodge(0.75),
            vjust = -0.4, size = 4.2, fontface = "bold") +
  scale_fill_manual(values = c(`raw p < 0.05` = "grey78", `FDR < 5%` = "#E76F51"), name = NULL) +
  scale_x_discrete(labels = c("Community type\n(Lactobacillus vs anaerobe)",
                              "Method\n(ring vs systemic)")) +
  labs(y = "number of host genes", x = NULL,
       title = "Which explanation fits host gene expression better?") +
  coord_cartesian(ylim = c(0, 2400)) +
  theme_bw(base_size = 13) +
  theme(legend.position = "bottom", axis.text.x = element_text(size = 12))
ggsave(here("group_project/slides_assets/fig3_headtohead.png"), fig3,
       width = 7.2, height = 5.4, dpi = 300, bg = "white")

cat("figures written\n")
