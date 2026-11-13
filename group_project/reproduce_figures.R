# Regenerate the six figures the deck embeds, using the exact plotting code
# from the two QC reports. Output: slides_assets/*.png at 1344x960.
suppressPackageStartupMessages({
  library(tidyverse)
  library(here)
  library(microViz)
  source(here("R/count_tables.R"))
})

dir.create(here("group_project/group_b_2/slides_assets"), showWarnings = FALSE)
out <- function(name) here("group_project/group_b_2/slides_assets", name)
save_fig <- function(p, name) ggsave(out(name), p, width = 7, height = 5, dpi = 192)

d <- read_counts("group_b")
ps_b <- as_counts("group_b")

## ---- priv_1 : 10 smallest microbiome libraries (Priv's Check 4) ----
libs <- colSums(d$species)
p1 <- tibble(sample_id = names(sort(libs)), reads = sort(libs)) |>
  mutate(lowest = row_number() <= 10) |>
  ggplot(aes(reorder(sample_id, reads), reads, fill = lowest)) +
  geom_col() +
  coord_flip() +
  scale_y_continuous(labels = scales::comma) +
  labs(x = NULL, y = "reads per sample",
       title = "10 smallest microbiome libraries",
       subtitle = "The three arms share one colour scale — no arm is systematically under-sequenced") +
  theme_bw(base_size = 10) +
  theme(legend.position = "none")
save_fig(p1, "priv_1.png")

## ---- priv_2 / priv_3 : Bray-Curtis PCoA by arm and by batch (Priv's Check 7) ----
o <- ps_b |>
  tax_transform("identity", rank = "species") |>
  dist_calc("bray") |>
  ord_calc("PCoA")
p2 <- ord_plot(o, colour = "arm") + theme_bw(base_size = 10) + ggtitle("coloured by arm")
p3 <- ord_plot(o, colour = "rnaseq_batch") + theme_bw(base_size = 10) + ggtitle("coloured by batch")
save_fig(p2, "priv_2.png")
save_fig(p3, "priv_3.png")

## ---- samwl_1 : library sizes, both tables (Samwl) ----
lib <- tibble(
  sample_id = colnames(d$species),
  species_reads = colSums(d$species),
  gene_reads = colSums(d$genes)
)
p4 <- lib |>
  pivot_longer(-sample_id, names_to = "table", values_to = "reads") |>
  ggplot(aes(x = reads, y = reorder(sample_id, reads), fill = table)) +
  geom_col(position = "dodge") +
  facet_wrap(~table, scales = "free_x") +
  labs(x = "reads per sample", y = NULL, title = "Library sizes, Study B") +
  theme(axis.text.y = element_blank(), axis.ticks.y = element_blank())
save_fig(p4, "samwl_1.png")

## ---- samwl_2 : feature prevalence histogram (Samwl) ----
prev <- rowSums(d$species > 0)
p5 <- tibble(prevalence = prev) |>
  ggplot(aes(x = prevalence)) +
  geom_histogram(bins = 40) +
  labs(x = "number of samples a species appears in", y = "species",
       title = paste(ncol(d$species), "samples,", nrow(d$species), "species"))
save_fig(p5, "samwl_2.png")

## ---- samwl_3 : Shannon index by arm (Samwl, exploratory) ----
shannon <- apply(d$species, 2, \(x) {
  p <- x / sum(x)
  p <- p[p > 0]
  -sum(p * log(p))
})
shan <- tibble(sample_id = names(shannon), shannon = shannon) |>
  left_join(d$metadata |> select(sample_id, arm), by = "sample_id")
p6 <- shan |>
  ggplot(aes(x = arm, y = shannon, fill = arm)) +
  geom_boxplot(alpha = 0.6, outlier.shape = NA) +
  geom_jitter(width = 0.15, size = 1.5, alpha = 0.7) +
  labs(y = "Shannon index", x = NULL,
       title = "Within-sample diversity by contraceptive method") +
  theme(legend.position = "none")
save_fig(p6, "samwl_3.png")

cat("figures written:\n")
for (f in c("priv_1", "priv_2", "priv_3", "samwl_1", "samwl_2", "samwl_3")) {
  cat(" ", f, file.size(out(paste0(f, ".png"))), "bytes\n")
}
