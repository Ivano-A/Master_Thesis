library(readr)
library(haven)
library(readxl)
library(writexl)
library(tidyverse)
library(ggthemes)
library(ggridges)
library(forcats)

# Line plot loop for all the provinces ####

file_paths <- c("Counts/data_for_descriptive/male_ON_descriptive.xlsx",
                "Counts/data_for_descriptive/male_QC_descriptive.xlsx",
                "Counts/data_for_descriptive/male_SK_descriptive.xlsx",
                "Counts/data_for_descriptive/male_MB_descriptive.xlsx",
                "Counts/data_for_descriptive/male_BC_descriptive.xlsx",
                "Counts/data_for_descriptive/male_AB_descriptive.xlsx",
                "Counts/data_for_descriptive/male_atlantic_descriptive.xlsx",
                "Counts/data_for_descriptive/female_SKMB_descriptive.xlsx",
                "Counts/data_for_descriptive/female_QC_descriptive.xlsx",
                "Counts/data_for_descriptive/female_ON_descriptive.xlsx",
                "Counts/data_for_descriptive/female_BC_descriptive.xlsx",
                "Counts/data_for_descriptive/female_atlantic_descriptive.xlsx",
                "Counts/data_for_descriptive/female_AB_descriptive.xlsx")

# Colour-blind pallet for the line graph
okabe_ito_colors <- c("#E69F00", "#56B4E9", "#009E73", "#F0E442", 
                      "#0072B2", "#D55E00", "#CC79A7", "#999999",
                      "#E6AB00", "#EEBAB4", "black", "#1F449C",
                      "#A8B6CC", "#F05039")

# Linetypes for the line graph
linetypes <- c("solid", "solid", "solid", "solid", "solid", "solid", "solid",
               "dashed", "dashed", "dashed", "dashed", "dashed", "dashed", "dashed")

period_line_list <- list()

for (file_path in file_paths) {
  # Read data
  data <- read_xlsx(file_path)
  # Get the max age so that I can graph the data to the proper age scale
  max_age.n <- max(data$age.n)
  
  if (max_age.n == 67) {
    
    plot_period_line <- data %>% 
      mutate(period_for_plot = as.factor(period_for_plot)) 
    
    period_plot <- ggplot(plot_period_line, aes(x = age.n, y = rate, colour = period_for_plot, 
                                                linetype = period_for_plot)) +
      geom_line(na.rm = TRUE) +
      scale_x_continuous(breaks = seq(17, 67, by = 5),  # Define the points where labels should appear
                         labels = c("15-19", "20-24", "25-29", "30-34", "35-39", 
                                    "40-44", "45-49", "50-54", "55-59", "60-64", 
                                    "65-84"),  # Custom labels for each break
                         limits = c(17, 67)  # Set limits to match your data to prevent extra grid lines
      ) +
      #scale_y_continuous(breaks = seq(0, max(plot_period_line$rate, na.rm = TRUE), by = 5)) +  # Adjust y-axis breaks
      labs(x = "Age",
           y = "Suicide Rate per 100,000",
           colour = "Period",
           linetype = "Period") +
      scale_color_manual(values = okabe_ito_colors) +
      scale_linetype_manual(values = linetypes) +
      theme_minimal() +
      theme(legend.title = element_text(face = "bold"),
            legend.position = "right",
            panel.grid.major = element_line(colour = "grey85"),
            panel.grid.minor = element_blank(),
            plot.background = element_rect(fill = "white"), 
      )
    
    period_line_list[[file_path]] <- period_plot 
  }
  
  # Line plot of each period, by age
  
  if (max_age.n == 72) {
    
    plot_period_line <- data %>% 
      mutate(period_for_plot = as.factor(period_for_plot)) 
    
    period_plot <- ggplot(plot_period_line, aes(x = age.n, y = rate, colour = period_for_plot, 
                                                linetype = period_for_plot)) +
      geom_line(na.rm = TRUE) +
      scale_x_continuous(breaks = seq(17, 72, by = 5),  # Define the points where labels should appear
                         labels = c("15-19", "20-24", "25-29", "30-34", "35-39", 
                                    "40-44", "45-49", "50-54", "55-59", "60-64", 
                                    "65-69", "70-84"),  # Custom labels for each break
                         limits = c(17, 72)  # Set limits to match your data to prevent extra grid lines
      ) +
      #scale_y_continuous(breaks = seq(0, max(plot_period_line$rate, na.rm = TRUE), by = 5)) +  # Adjust y-axis breaks
      labs(x = "Age",
           y = "Suicide Rate per 100,000",
           colour = "Period",
           linetype = "Period") +
      scale_color_manual(values = okabe_ito_colors) +
      scale_linetype_manual(values = linetypes) +
      theme_minimal() +
      theme(legend.title = element_text(face = "bold"),
            legend.position = "right",
            panel.grid.major = element_line(colour = "grey85"),
            panel.grid.minor = element_blank(),
            plot.background = element_rect(fill = "white"), 
      )
    
    period_line_list[[file_path]] <- period_plot 
  }
  
  if (max_age.n == 77) {
    
    plot_period_line <- data %>% 
      mutate(period_for_plot = as.factor(period_for_plot)) 
    
    period_plot <- ggplot(plot_period_line, aes(x = age.n, y = rate, colour = period_for_plot, 
                                                linetype = period_for_plot)) +
      geom_line(na.rm = TRUE) +
      scale_x_continuous(breaks = seq(17, 77, by = 5),  # Define the points where labels should appear
                         labels = c("15-19", "20-24", "25-29", "30-34", "35-39", 
                                    "40-44", "45-49", "50-54", "55-59", "60-64", 
                                    "65-69", "70-74", "75-84"),  # Custom labels for each break
                         limits = c(17, 77)  # Set limits to match your data to prevent extra grid lines
      ) +
      #scale_y_continuous(breaks = seq(0, max(plot_period_line$rate, na.rm = TRUE), by = 5)) +  # Adjust y-axis breaks
      labs(x = "Age",
           y = "Suicide Rate per 100,000",
           colour = "Period",
           linetype = "Period") +
      scale_color_manual(values = okabe_ito_colors) +
      scale_linetype_manual(values = linetypes) +
      theme_minimal() +
      theme(legend.title = element_text(face = "bold"),
            legend.position = "right",
            panel.grid.major = element_line(colour = "grey85"),
            panel.grid.minor = element_blank(),
            plot.background = element_rect(fill = "white"), 
      )
    
    period_line_list[[file_path]] <- period_plot 
  }
  
  if (max_age.n == 82) {
    
    plot_period_line <- data %>% 
      mutate(period_for_plot = as.factor(period_for_plot)) 
    
    period_plot <- ggplot(plot_period_line, aes(x = age.n, y = rate, colour = period_for_plot, 
                                                linetype = period_for_plot)) +
      geom_line(na.rm = TRUE) +
      scale_x_continuous(breaks = seq(17, 82, by = 5),  # Define the points where labels should appear
                         labels = c("15-19", "20-24", "25-29", "30-34", "35-39", 
                                    "40-44", "45-49", "50-54", "55-59", "60-64", 
                                    "65-69", "70-74", "75-79", "80-84"),  # Custom labels for each break
                         limits = c(17, 82)  # Set limits to match your data to prevent extra grid lines
      ) +
      #scale_y_continuous(breaks = seq(0, max(plot_period_line$rate, na.rm = TRUE), by = 5)) +  # Adjust y-axis breaks
      labs(x = "Age",
           y = "Suicide Rate per 100,000",
           colour = "Period",
           linetype = "Period") +
      scale_color_manual(values = okabe_ito_colors) +
      scale_linetype_manual(values = linetypes) +
      theme_minimal() +
      theme(legend.title = element_text(face = "bold"),
            legend.position = "right",
            panel.grid.major = element_line(colour = "grey85"),
            panel.grid.minor = element_blank(),
            plot.background = element_rect(fill = "white"), 
      )
    
    period_line_list[[file_path]] <- period_plot 
    
  }
}

period_line_list

pdf("period_line_plots_v4_final.pdf", width = 11, height = 8)  # Adjust width and height as needed

for (plot in period_line_list) {
  print(plot)
}

dev.off()

library(readr)
library(haven)
library(readxl)
library(writexl)
library(tidyverse)
library(ggthemes)
library(ggridges)
library(forcats)

# By period loop ####

file_paths <- c("Counts/data_for_descriptive/male_ON_descriptive.xlsx",
                "Counts/data_for_descriptive/male_QC_descriptive.xlsx",
                "Counts/data_for_descriptive/male_SK_descriptive.xlsx",
                "Counts/data_for_descriptive/male_MB_descriptive.xlsx",
                "Counts/data_for_descriptive/male_BC_descriptive.xlsx",
                "Counts/data_for_descriptive/male_AB_descriptive.xlsx",
                "Counts/data_for_descriptive/male_atlantic_descriptive.xlsx",
                "Counts/data_for_descriptive/female_SKMB_descriptive.xlsx",
                "Counts/data_for_descriptive/female_QC_descriptive.xlsx",
                "Counts/data_for_descriptive/female_ON_descriptive.xlsx",
                "Counts/data_for_descriptive/female_BC_descriptive.xlsx",
                "Counts/data_for_descriptive/female_atlantic_descriptive.xlsx",
                "Counts/data_for_descriptive/female_AB_descriptive.xlsx")

faceted_plot <- list()

for (file_path in file_paths) {
  
  data <- read_xlsx(file_path)
  
  faceted_period_plot <- ggplot(data, aes(x = as.numeric(age.n), y = rate), na.rm = TRUE) +
    geom_line() +
    geom_point() +
    facet_wrap(~period_for_plot) +#, scales = "free") +
    scale_x_continuous(breaks = seq(17, 85, by = 10),
                       labels = c("15-9", "25-9",
                                  "35-9", "45-9",
                                  "55-9", "65-9",
                                  "75-9")) +
    xlab("Age") +
    ylab("Suicide Rate per 100,000") +
    scale_colour_colorblind() +
    theme_minimal() +
    theme(
      legend.position = c(0.6, 0.1),
      legend.box.just = c("right", "bottom"),
      legend.title = element_blank(),
      # panel.background = element_rect(fill = "white"), # Plot area background
      #plot.background = element_rect(fill = "white"), 
      #panel.border = element_blank(),
      #panel.grid.major = element_blank(),
      panel.grid.minor = element_blank(),
      #panel.background = element_rect(fill = "white"),
      panel.border = element_rect(colour = "black", fill = NA),
      strip.background = element_rect(colour = "black", fill = "grey85")
    ) + 
    ggtitle(file_path)
  
  faceted_plot[[file_path]] <- faceted_period_plot
  
}

faceted_plot

pdf("period_faceted_plot_for_appendix.pdf", width = 11, height = 8)  # Adjust width and height as needed

for (plot in faceted_plot) {
  print(plot)
}

dev.off()

# Bar Chart of youth age groups 15-24 ####

library(readr)
library(haven)
library(readxl)
library(writexl)
library(tidyverse)
library(ggthemes)
library(ggridges)
library(forcats)

file_paths <- c("Counts/data_for_descriptive/male_ON_descriptive.xlsx",
                "Counts/data_for_descriptive/male_QC_descriptive.xlsx",
                "Counts/data_for_descriptive/male_SK_descriptive.xlsx",
                "Counts/data_for_descriptive/male_MB_descriptive.xlsx",
                "Counts/data_for_descriptive/male_BC_descriptive.xlsx",
                "Counts/data_for_descriptive/male_AB_descriptive.xlsx",
                "Counts/data_for_descriptive/male_atlantic_descriptive.xlsx",
                "Counts/data_for_descriptive/female_SKMB_descriptive.xlsx",
                "Counts/data_for_descriptive/female_QC_descriptive.xlsx",
                "Counts/data_for_descriptive/female_ON_descriptive.xlsx",
                "Counts/data_for_descriptive/female_BC_descriptive.xlsx",
                "Counts/data_for_descriptive/female_atlantic_descriptive.xlsx",
                "Counts/data_for_descriptive/female_AB_descriptive.xlsx")

bar_chart_list <- list()

for (file_path in file_paths) {
  
  data <- read_xlsx(file_path)
  
  filtered_youth_groups <- data %>% 
    select(rate, period_for_plot, age.n, yr.n) %>% 
    filter(age.n %in% c(17, 22))
  
  base_name <- str_remove(basename(file_path),"_descriptive.xlsx")
  
  bar_chart <- ggplot(filtered_youth_groups, aes(x = period_for_plot, y = rate, fill = factor(age.n))) +
    geom_bar(stat = "identity", position = "dodge") +  # Use "dodge" to place bars side by side
    labs(x = "Period", y = "Suicide Rate per 100,000", fill = "Age Group",
         title = base_name) +
    geom_text(aes(label = round(rate, 1)),    # Add rounded rate values as labels
              position = position_dodge(width = 0.9), # Align text with dodged bars
              vjust = -0.5,                   # Position above the bars
              size = 3) +                     # Adjust text size
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          plot.title = element_text(hjust = 0.5, size = 14, face = "bold")) +
    scale_fill_manual(
      values = c("#1f78b4", "#33a02c"),
      labels = c("15-19", "20-24")
    ) +
  ylim(0, max(filtered_youth_groups$rate) * 1.2)  # Increase upper limit by 20%
  
  bar_chart_list[[file_path]] <- bar_chart
  
}

pdf("bar_chart_youth_groups_with_rates.pdf", width = 11, height = 8)  # Adjust width and height as needed


for (plot in bar_chart_list) {
  print(plot)
}

dev.off()

# Saving for presentation in thesis, 3 per page and one 4 per page

library(ggpubr)

# Assuming bar_chart_list contains 13 ggplot objects
n_plots <- length(bar_chart_list)
plots_per_page <- 3

# Create groupings of 3 per page, except the last one
grouped_plots <- split(bar_chart_list, ceiling(seq_along(bar_chart_list) / plots_per_page))

# Ensure the last page has 4 plots
if (length(grouped_plots) > 1) {
  last_page <- tail(grouped_plots, 1)[[1]]
  if (length(last_page) < 4 && length(grouped_plots) > 1) {
    grouped_plots[[length(grouped_plots) - 1]] <- c(grouped_plots[[length(grouped_plots) - 1]], last_page)
    grouped_plots <- grouped_plots[-length(grouped_plots)]
  }
}

# Save each page
for (i in seq_along(grouped_plots)) {
  p <- ggarrange(plotlist = grouped_plots[[i]], ncol = 1, nrow = ifelse(i == length(grouped_plots), 4, 3))
  ggsave(filename = paste0("bar_charts_page_", i, ".jpg"), plot = p, width = 8, height = 10, dpi = 300)
}


# Plotting by cohort ####

library(readr)
library(haven)
library(readxl)
library(writexl)
library(tidyverse)
library(ggthemes)
library(ggridges)
library(forcats)

# Some of the datasets needed to have the cohort variable added, filtering for only the
# 1935 to 1990 cohorts
# Overlaying every 

# Male cohort rates: 

QC_M <- read_xlsx("Counts/data_for_descriptive/male_QC_descriptive.xlsx") %>% 
  mutate(province = "Québec") %>% 
  filter(cohort %in% c(1935:1990)) %>%  
  rename(cohort_low = cohort) %>% 
  mutate(cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  )

ON_M <- read_xlsx("Counts/data_for_descriptive/male_ON_descriptive.xlsx") %>% 
  mutate(province = "Ontario") %>% 
  filter(cohort %in% c(1935:1990)) %>%  
  rename(cohort_low = cohort) %>% 
  mutate(cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  )

BC_M <- read_xlsx("Counts/data_for_descriptive/male_BC_descriptive.xlsx") %>% 
  mutate(province = "British Columbia") %>% 
  filter(cohort %in% c(1935:1990)) %>%  
  rename(cohort_low = cohort) %>% 
  mutate(cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  )

SK_M <- read_xlsx("Counts/data_for_descriptive/male_SK_descriptive.xlsx") %>% 
  mutate(province = "Saskatchewan",
         cohort_low = yr.n - age.n,
         cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  ) %>% 
  filter(cohort_low %in% c(1935:1990))

MB_M <- read_xlsx("Counts/data_for_descriptive/male_MB_descriptive.xlsx") %>% 
  mutate(province = "Manitoba",
         cohort_low = yr.n - age.n,
         cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  ) %>% 
  filter(cohort_low %in% c(1935:1990))

ATL_M <- read_xlsx("Counts/data_for_descriptive/male_atlantic_descriptive.xlsx") %>% 
  mutate(province = "Atlantic Canada",
         cohort_low = yr.n - age.n,
         cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  ) %>% 
  filter(cohort_low %in% c(1935:1990))

AB_M <- read_xlsx("Counts/data_for_descriptive/male_AB_descriptive.xlsx") %>% 
  mutate(province = "Alberta",
         cohort_low = yr.n - age.n,
         cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  ) %>% 
  filter(cohort_low %in% c(1935:1990))

Male_cohort_plot <- ggplot() +
  geom_line(data = QC_M, aes(x = age.n, y = rate, colour = province)) +
  geom_line(data = ON_M, aes(x = age.n, y = rate, colour = province)) +
  geom_line(data = BC_M, aes(x = age.n, y = rate, colour = province)) +
  geom_line(data = MB_M, aes(x = age.n, y = rate, colour = province)) +
  geom_line(data = SK_M, aes(x = age.n, y = rate, colour = province)) +
  geom_line(data = AB_M, aes(x = age.n, y = rate, colour = province)) +
  geom_line(data = ATL_M, aes(x = age.n, y = rate, colour = province)) +
  facet_wrap(~cohort) +
  scale_x_continuous(breaks = seq(17, 85, by = 10),
                     labels = c("15-19", "25-29",
                                "35-39", "45-49",
                                "55-59", "65-69",
                                "75-79")) +
  xlab("Age") +
  ylab("Suicide Rate per 100,000") +
  #ggtitle("Suicide Rates by Cohort, Male") +
  labs(colour = "Provinces") +
  theme_minimal() +
  theme(
    panel.grid.minor = element_blank(),
    panel.border = element_rect(colour = "black", fill = NA),
    strip.background = element_rect(colour = "black", fill = "grey85"),
    axis.text.x = element_text(angle = 45, hjust = 0.8, vjust = 1)
  ) +
  scale_color_colorblind() 



Male_cohort_plot

# Female cohort rates

QC_F <- read_xlsx("Counts/data_for_descriptive/female_QC_descriptive.xlsx") %>% 
  mutate(province = "Québec",
         cohort_low = yr.n - age.n,
         cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  ) %>% 
  filter(cohort_low %in% c(1935:1990))

ON_F <- read_xlsx("Counts/data_for_descriptive/female_ON_descriptive.xlsx") %>% 
  mutate(province = "Ontario",
         cohort_low = yr.n - age.n,
         cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  ) %>% 
  filter(cohort_low %in% c(1935:1990))

BC_F <- read_xlsx("Counts/data_for_descriptive/female_BC_descriptive.xlsx") %>% 
  mutate(province = "British Columbia",
         cohort_low = yr.n - age.n,
         cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  ) %>% 
  filter(cohort_low %in% c(1935:1990))

SKMB_F <- read_xlsx("Counts/data_for_descriptive/female_SKMB_descriptive.xlsx") %>% 
  mutate(province = "Manitoba and Saskatchewan",
         cohort_low = yr.n - age.n,
         cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  ) %>% 
  filter(cohort_low %in% c(1935:1990))

AB_F <- read_xlsx("Counts/data_for_descriptive/female_AB_descriptive.xlsx") %>% 
  mutate(province = "Alberta",
         cohort_low = yr.n - age.n,
         cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  ) %>% 
  filter(cohort_low %in% c(1935:1990))

ATL_F <- read_xlsx("Counts/data_for_descriptive/female_atlantic_descriptive.xlsx") %>% 
  mutate(province = "Atlantic Canada",
         cohort_low = yr.n - age.n,
         cohort_up = cohort_low + 4,
         cohort = paste0(cohort_low, "-", cohort_up)
  ) %>% 
  filter(cohort_low %in% c(1935:1990))


Female_cohort_plot <- ggplot() +
  geom_line(data = QC_F, aes(x = age.n, y = rate, colour = province)) +
  geom_line(data = ON_F, aes(x = age.n, y = rate, colour = province)) +
  geom_line(data = BC_F, aes(x = age.n, y = rate, colour = province)) +
  geom_line(data = SKMB_F, aes(x = age.n, y = rate, colour = province)) +
  geom_line(data = AB_F, aes(x = age.n, y = rate, colour = province)) +
  geom_line(data = ATL_F, aes(x = age.n, y = rate, colour = province)) +
  facet_wrap(~cohort) +
  scale_x_continuous(breaks = seq(17, 85, by = 10),
                     labels = c("15-19", "25-29",
                                "35-39", "45-49",
                                "55-59", "65-69",
                                "75-79")) +
  xlab("Age") +
  ylab("Suicide Rate per 100,000") +
  #ggtitle("Suicide Rates by Cohort, Female") +
  labs(colour = "Provinces") +
  theme_minimal() +
  theme(
    panel.grid.minor = element_blank(),
    panel.border = element_rect(colour = "black", fill = NA),
    strip.background = element_rect(colour = "black", fill = "grey85"),
    axis.text.x = element_text(angle = 45, hjust = 0.8, vjust = 1)
  ) +
  scale_color_colorblind() 


Female_cohort_plot

# Saving the cohort plots

pdf("Male_cohort_plots.pdf", width = 11, height = 8)  # Adjust width and height as needed

print(Male_cohort_plot)

dev.off()

pdf("female_cohort_plots.pdf", width = 11, height = 8)  # Adjust width and height as needed

print(Female_cohort_plot)

dev.off()


# Sex-ratios ####

library(readr)
library(haven)
library(readxl)
library(writexl)
library(tidyverse)
library(ggthemes)
library(ggridges)
library(forcats)

file_paths <- c("Counts/data_for_descriptive/male_ON_descriptive.xlsx",
                "Counts/data_for_descriptive/male_QC_descriptive.xlsx",
                "Counts/data_for_descriptive/male_SK_descriptive.xlsx",
                "Counts/data_for_descriptive/male_MB_descriptive.xlsx",
                "Counts/data_for_descriptive/male_BC_descriptive.xlsx",
                "Counts/data_for_descriptive/male_AB_descriptive.xlsx",
                "Counts/data_for_descriptive/male_atlantic_descriptive.xlsx",
                "Counts/data_for_descriptive/female_SKMB_descriptive.xlsx",
                "Counts/data_for_descriptive/female_QC_descriptive.xlsx",
                "Counts/data_for_descriptive/female_ON_descriptive.xlsx",
                "Counts/data_for_descriptive/female_BC_descriptive.xlsx",
                "Counts/data_for_descriptive/female_atlantic_descriptive.xlsx",
                "Counts/data_for_descriptive/female_AB_descriptive.xlsx")



# Define a function to read and rename columns based on the file path
read_and_rename <- function(file_path) {
  # Read the dataset
  df <- read_excel(file_path)
  
  # Determine sex and province from the file path
  sex <- ifelse(grepl("female", file_path), "female", "male")
  province <- gsub(".*_(\\w{2,8})_.*", "\\1", file_path)
  
  # Rename the columns based on sex and province
  df <- df %>%
    rename_with(~ paste0(., "_", sex, "_", province), 
                .cols = c("rate", "count", "tot_p"))  # Rename if these columns exist
  
  return(df)
}

# Initialize an empty dataframe
combined_data <- NULL

# Iterate over the file paths and combine the datasets using a full join
for (file in file_paths) {
  df <- read_and_rename(file) %>% 
    select(-age, -period_for_plot)
  
  if (is.null(combined_data)) {
    # For the first file, assign it directly to combined_data
    combined_data <- df
  } else {
    # For subsequent files, join with the existing combined dataset on age.n and yr.n
    combined_data <- full_join(combined_data, df, by = c("age.n", "yr.n"))
  }
}


# Filtering the merged dataset as some provinces had cohort or year5 variables, and rounding values to 2 decimal places
filtered_df <- combined_data %>% 
  select(-starts_with("cohort"), -starts_with("year")) %>% 
  mutate(round(across(starts_with("rate")), 2))


# Creating a dataset with every province having age groups of 65+ ####

###_____NEED filtered_df from above______###

library(tidyverse)
library(readr)
library(writexl)

file_paths <- c("Counts/data_for_descriptive/male_ON_descriptive.xlsx",
                "Counts/data_for_descriptive/male_QC_descriptive.xlsx",
                "Counts/data_for_descriptive/male_SK_descriptive.xlsx",
                "Counts/data_for_descriptive/male_MB_descriptive.xlsx",
                "Counts/data_for_descriptive/male_BC_descriptive.xlsx",
                "Counts/data_for_descriptive/male_AB_descriptive.xlsx",
                "Counts/data_for_descriptive/male_atlantic_descriptive.xlsx",
                #"Counts/data_for_descriptive/female_SKMB_descriptive.xlsx",
                "Counts/data_for_descriptive/female_QC_descriptive.xlsx",
                "Counts/data_for_descriptive/female_ON_descriptive.xlsx",
                "Counts/data_for_descriptive/female_BC_descriptive.xlsx",
                #"Counts/data_for_descriptive/female_atlantic_descriptive.xlsx",
                "Counts/data_for_descriptive/female_AB_descriptive.xlsx")

age_group_67_plus <- list()

for (file_path in file_paths) {
  
  data <- read_xlsx(file_path)
  
  # Create dynamic column names outside of summarise
  count_col_name <- paste0("count_", basename(str_remove(file_path, "_descriptive.xlsx")))
  tot_p_col_name <- paste0("tot_p_", basename(str_remove(file_path, "_descriptive.xlsx")))
  rate_col_name <- paste0("rate_", basename(str_remove(file_path, "_descriptive.xlsx")))
  
  summing_rows <- data %>% 
    filter(age.n >= 67) %>% 
    group_by(yr.n) %>% 
    summarise(
      !!count_col_name := sum(count),
      !!tot_p_col_name := sum(tot_p)
    ) %>% 
    mutate(!!rate_col_name := .data[[count_col_name]] / .data[[tot_p_col_name]] * 100000,
           age.n = 67
    ) %>% 
    relocate(age.n)
  
  # Store the result in the list
  age_group_67_plus[[file_path]] <- summing_rows
  
}

combined_df_67 <- reduce(age_group_67_plus, full_join, by = c("age.n", "yr.n")) %>% 
  as.data.frame() 


# Filter to get only the rows that are age.n 62 or less and the columns that needed
# Combine values of 67 plus combined to match female SKMB and female atlantic
filtered_df_to_remove_rows <- filtered_df %>% 
  filter(age.n <= 62) %>% 
  relocate(age.n) %>% 
  select(age.n, yr.n,
         count_male_ON:rate_male_atlantic, 
         count_female_QC:rate_female_BC, 
         count_female_AB:rate_female_AB) 

# Row binding the values of the combined 67+ rows and the filtered age.n = 62 above
combined_62_67 <- rbind(filtered_df_to_remove_rows, combined_df_67) 

# Adding back the female Atlantic and female SKMB
fem_SKMB_atlantic <- filtered_df %>% 
  select(count_female_atlantic:rate_female_atlantic, 
         count_female_SKMB:rate_female_SKMB) %>% 
  slice(1:154)

# Combining fem Atlantic and SKMB to combined_62_67
complete_67_plus_df <- cbind(combined_62_67, fem_SKMB_atlantic) 

complete_67_plus_df <- complete_67_plus_df %>% 
  relocate(count_female_SKMB:rate_female_SKMB, .after = rate_female_AB) %>% 
  relocate(count_male_AB:rate_male_AB, .after = rate_male_QC)

write_xlsx(complete_67_plus_df, path = "counts_rates_67_plus.xlsx")

# Calculating sex-ratios for 15-67+
sex_ratio_67_plus <- complete_67_plus_df %>% 
  mutate(
    # Calculate sex ratios for each province
    sex_ratio_QC = ifelse(rate_female_QC == 0 | is.na(rate_female_QC), NA, rate_male_QC / rate_female_QC),
    sex_ratio_ON = ifelse(rate_female_ON == 0 | is.na(rate_female_ON), NA, rate_male_ON / rate_female_ON),
    sex_ratio_MB = ifelse(rate_female_SKMB == 0 | is.na(rate_female_SKMB), NA, rate_male_MB / rate_female_SKMB),
    sex_ratio_SK = ifelse(rate_female_SKMB == 0 | is.na(rate_female_SKMB), NA, rate_male_SK / rate_female_SKMB),
    sex_ratio_BC = ifelse(rate_female_BC == 0 | is.na(rate_female_BC), NA, rate_male_BC / rate_female_BC),
    sex_ratio_AB = ifelse(rate_female_AB == 0 | is.na(rate_female_AB), NA, rate_male_AB / rate_female_AB),
    sex_ratio_atlantic = ifelse(rate_female_atlantic == 0 | is.na(rate_female_atlantic), NA, rate_male_atlantic / rate_female_atlantic)
  ) %>% 
  select(age.n, yr.n, starts_with("sex_ratio"))

write_xlsx(sex_ratio_67_plus, path = "sex_ratio_all_to_67_plus.xlsx")

# Sex ratio by age
sex_ratio_by_age <- sex_ratio_67_plus %>% 
  group_by(age.n) %>% 
  summarise(across(starts_with("sex"), ~ mean(.x, na.rm = TRUE))) %>% 
  rename(age_year = age.n) # matching the column name with the year df

# Sex ratio by year
sex_ratio_by_year <- sex_ratio_67_plus %>% 
  group_by(yr.n) %>% 
  summarise(across(starts_with("sex"), ~ mean(.x, na.rm = TRUE))) %>%  
  rename(age_year = yr.n)


# Combining the dfs of sex-ratio by age and year
sex_ratio_by_age_and_year_combined <- rbind(sex_ratio_by_age, sex_ratio_by_year)

write_xlsx(sex_ratio_by_age_and_year_combined, path = "sex_ratio_age_and_year_combined.xlsx")

# bar chart for age sex-ratios
age_labels <- c("15-19", "20-24", "25-29", "30-34", "35-39", "40-44", 
                "45-49", "50-54", "55-59", "60-64", "65-84")

province_labels <- c("Alberta", "Atlantic", "British Columbia", "Manitoba", 
                     "Ontario", "Québec", "Saskatchewan")

sex_ratio_long_age <- sex_ratio_by_age %>%
  pivot_longer(cols = starts_with("sex_ratio"), 
               names_to = "province", 
               values_to = "sex_ratio")

avg_sex_ratio_by_age_and_province <- ggplot(sex_ratio_long, aes(x = as.factor(age_year), y = sex_ratio, fill = as.factor(province))) +
  geom_bar(stat = "identity", position = "dodge") +
  labs(title = "Average Sex-ratios by Age Group and Province",
       x = "Age Group", 
       y = "Sex Ratio (Male/Female)") +
  scale_x_discrete(labels = age_labels) +
  scale_fill_viridis_d(name = "Province",
                       labels = province_labels) + # Color bars by year
  theme_minimal()

# Bar chart for average sex-ratio by period

# Bar chart labels
period_labels <- c("1950-1954", "1955-1959", "1960-1964", "1965-1969", "1970-1974", 
                   "1975-1979", "1980-1984", "1985-1989", "1990-1994", "1995-1999", 
                   "2000-2004", "2005-2009", "2010-2014", "2015-2019")

province_labels <- c("Alberta", "Atlantic", "British Columbia", "Manitoba", 
                     "Ontario", "Québec", "Saskatchewan")

sex_ratio_long_year <- sex_ratio_by_year %>%
  pivot_longer(cols = starts_with("sex_ratio"), 
               names_to = "province", 
               values_to = "sex_ratio")

avg_sex_ratio_by_year_and_province <- ggplot(sex_ratio_long_year, aes(x = as.factor(age_year), y = sex_ratio, fill = as.factor(province))) +
  geom_bar(stat = "identity", position = "dodge") +
  labs(title = "Average Sex-ratios by Period and Province",
       x = "Period", 
       y = "Sex Ratio (Male/Female)") +
  scale_x_discrete(labels = period_labels) +
  scale_fill_viridis_d(name = "Province",
                       labels = province_labels) + # Color bars by year
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

avg_sex_ratio_by_year_and_province

pdf("average_sex_ratio_for_age_and_period.pdf", width = 11, height = 8) 

print(avg_sex_ratio_by_age_and_province)
print(avg_sex_ratio_by_year_and_province)

dev.off()

# Bar charts sex-ratios ####

library(readr)
library(haven)
library(readxl)
library(writexl)
library(tidyverse)
library(ggthemes)
library(ggridges)
library(forcats)

# Load the data
sex_ratio_67_plus <- read_xlsx("sex_ratio_all_to_67_plus.xlsx")

# Bar Chart labels
age_labels <- c("15-19", "20-24", "25-29", "30-34", "35-39", "40-44", 
                "45-49", "50-54", "55-59", "60-64", "65-84")

period_labels <- c("1950-1954", "1955-1959", "1960-1964", "1965-1969", "1970-1974", 
                   "1975-1979", "1980-1984", "1985-1989", "1990-1994", "1995-1999", 
                   "2000-2004", "2005-2009", "2010-2014", "2015-2019")


# Quebec
QC_average_sex_Ratio <- sex_ratio_67_plus %>% 
  group_by(age.n) %>% 
  summarise(mean_sex_ratios_QC = mean(sex_ratio_QC, na.rm = TRUE))

QC_sex_ratio_bar_chart <- ggplot() +
  geom_bar(data = sex_ratio_67_plus, aes(x = as.factor(age.n), y = sex_ratio_QC, fill = as.factor(yr.n)), 
           stat = "identity", position = "dodge") +  # Bar plot showing sex ratios by age group and year
  # geom_line(data = QC_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_QC, group = 1), 
  #           color = "red", size = 1.5) +  # Line for average sex ratios across age groups 
  #geom_smooth(data = QC_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_QC, group = 1), 
  #method = "loess", color = "red", size = 1.5, se = FALSE) +
  labs(#title = "Sex-ratios by Age Group and Period in Québec", 
    x = "Age", 
    y = "Sex Ratio (Male/Female)") +
  geom_hline(yintercept = 1, linetype = "dashed", color = "red", size = 0.5) +  # Horizontal line at y = 1
  scale_x_discrete(labels = age_labels) +
  theme_minimal() +
  scale_fill_viridis_d(name = "Period",
                       labels = period_labels) # Color bars by year

QC_sex_ratio_bar_chart

# Bar Chart Manitoba

# line for bar chart showing average across each age group
MB_average_sex_Ratio <- sex_ratio_67_plus %>% 
  group_by(age.n) %>% 
  summarise(mean_sex_ratios_MB = mean(sex_ratio_MB, na.rm = TRUE))


MB_sex_ratio_bar_chart <- ggplot() +
  geom_bar(data = sex_ratio_67_plus, aes(x = as.factor(age.n), y = sex_ratio_MB, fill = as.factor(yr.n)), 
           stat = "identity", position = "dodge") +  # Bar plot showing sex ratios by age group and year
  #geom_line(data = MB_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_MB, group = 1), 
  #color = "red", size = 1.5) +
  # geom_smooth(data = MB_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_MB, group = 1), 
  #             method = "loess", color = "red", linewidth = 1.5, se = FALSE) +
  geom_hline(yintercept = 1, linetype = "dashed", color = "red", size = 0.5) +  # Horizontal line at y = 1
  # Line for average sex ratios across age groups 
  labs(#title = "Sex Ratios by Age Group and Year in Manitoba", 
    x = "Age", 
    y = "Sex Ratio (Male/Female)") +
  scale_x_discrete(labels = age_labels) +
  theme_minimal() +
  scale_fill_viridis_d(name = "Period",
                       labels = period_labels)  # Color bars by year

# Ontario

# line for bar chart showing average across each age group
ON_average_sex_Ratio <- sex_ratio_67_plus %>% 
  group_by(age.n) %>% 
  summarise(mean_sex_ratios_ON = mean(sex_ratio_ON, na.rm = TRUE))


ON_sex_ratio_bar_chart <- ggplot() +
  geom_bar(data = sex_ratio_67_plus, aes(x = as.factor(age.n), y = sex_ratio_ON, fill = as.factor(yr.n)), 
           stat = "identity", position = "dodge") +  # Bar plot showing sex ratios by age group and year
  #geom_line(data = MB_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_MB, group = 1), 
  #color = "red", size = 1.5) +
  # geom_smooth(data = ON_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_ON, group = 1), 
  #             method = "loess", color = "red", linewidth = 1.5, se = FALSE) +
  geom_hline(yintercept = 1, linetype = "dashed", color = "red", size = 0.5) +  # Horizontal line at y = 1
  # Line for average sex ratios across age groups 
  labs(#title = "Sex Ratios by Age Group and Year in Ontario", 
    x = "Age", 
    y = "Sex Ratio (Male/Female)") +
  scale_x_discrete(labels = age_labels) +
  theme_minimal() +
  scale_fill_viridis_d(name = "Period",
                       labels = period_labels)  # Color bars by year

ON_sex_ratio_bar_chart

# British Columbia

# line for bar chart showing average across each age group
BC_average_sex_Ratio <- sex_ratio_67_plus %>% 
  group_by(age.n) %>% 
  summarise(mean_sex_ratios_BC = mean(sex_ratio_BC, na.rm = TRUE))


BC_sex_ratio_bar_chart <- ggplot() +
  geom_bar(data = sex_ratio_67_plus, aes(x = as.factor(age.n), y = sex_ratio_BC, fill = as.factor(yr.n)), 
           stat = "identity", position = "dodge") +  # Bar plot showing sex ratios by age group and year
  #geom_line(data = MB_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_MB, group = 1), 
  #color = "red", size = 1.5) +
  # geom_smooth(data = BC_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_BC, group = 1), 
  #             method = "loess", color = "red", linewidth = 1.5, se = FALSE) +
  geom_hline(yintercept = 1, linetype = "dashed", color = "red", size = 0.5) +  # Horizontal line at y = 1
  # Line for average sex ratios across age groups 
  labs(#title = "Sex Ratios by Age Group and Year in British Columbia", 
    x = "Age", 
    y = "Sex Ratio (Male/Female)") +
  scale_x_discrete(labels = age_labels) +
  theme_minimal() +
  scale_fill_viridis_d(name = "Period",
                       labels = period_labels)  # Color bars by year

BC_sex_ratio_bar_chart

# Alberta

# line for bar chart showing average across each age group
AB_average_sex_Ratio <- sex_ratio_67_plus %>% 
  group_by(age.n) %>% 
  summarise(mean_sex_ratios_AB = mean(sex_ratio_AB, na.rm = TRUE))


AB_sex_ratio_bar_chart <- ggplot() +
  geom_bar(data = sex_ratio_67_plus, aes(x = as.factor(age.n), y = sex_ratio_AB, fill = as.factor(yr.n)), 
           stat = "identity", position = "dodge") +  # Bar plot showing sex ratios by age group and year
  #geom_line(data = MB_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_MB, group = 1), 
  #color = "red", size = 1.5) +
  # geom_smooth(data = AB_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_AB, group = 1), 
  #             method = "loess", color = "red", linewidth = 1.5, se = FALSE) +
  geom_hline(yintercept = 1, linetype = "dashed", color = "red", size = 0.5) +  # Horizontal line at y = 1
  # Line for average sex ratios across age groups 
  labs(#title = "Sex Ratios by Age Group and Year in Alberta", 
    x = "Age", 
    y = "Sex Ratio (Male/Female)") +
  scale_x_discrete(labels = age_labels) +
  theme_minimal() +
  scale_fill_viridis_d(name = "Period",
                       labels = period_labels)  # Color bars by year

AB_sex_ratio_bar_chart

# Saskatchewan

# line for bar chart showing average across each age group
SK_average_sex_Ratio <- sex_ratio_67_plus %>% 
  group_by(age.n) %>% 
  summarise(mean_sex_ratios_SK = mean(sex_ratio_SK, na.rm = TRUE))


SK_sex_ratio_bar_chart <- ggplot() +
  geom_bar(data = sex_ratio_67_plus, aes(x = as.factor(age.n), y = sex_ratio_SK, fill = as.factor(yr.n)), 
           stat = "identity", position = "dodge") +  # Bar plot showing sex ratios by age group and year
  #geom_line(data = MB_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_MB, group = 1), 
  #color = "red", size = 1.5) +
  # geom_smooth(data = SK_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_SK, group = 1), 
  #             method = "loess", color = "red", linewidth = 1.5, se = FALSE) +
  geom_hline(yintercept = 1, linetype = "dashed", color = "red", size = 0.5) +  # Horizontal line at y = 1
  # Line for average sex ratios across age groups 
  labs(#title = "Sex Ratios by Age Group and Year in Saskatchewan", 
    x = "Age", 
    y = "Sex Ratio (Male/Female)") +
  scale_x_discrete(labels = age_labels) +
  theme_minimal() +
  scale_fill_viridis_d(name = "Period",
                       labels = period_labels)  # Color bars by year

SK_sex_ratio_bar_chart

# atlantic

# line for bar chart showing average across each age group
atlantic_average_sex_Ratio <- sex_ratio_67_plus %>% 
  group_by(age.n) %>% 
  summarise(mean_sex_ratios_atlantic = mean(sex_ratio_atlantic, na.rm = TRUE))


atlantic_sex_ratio_bar_chart <- ggplot() +
  geom_bar(data = sex_ratio_67_plus, aes(x = as.factor(age.n), y = sex_ratio_atlantic, fill = as.factor(yr.n)), 
           stat = "identity", position = "dodge") +  # Bar plot showing sex ratios by age group and year
  #geom_line(data = MB_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_MB, group = 1), 
  #color = "red", size = 1.5) +
  # geom_smooth(data = atlantic_average_sex_Ratio, aes(x = as.factor(age.n), y = mean_sex_ratios_atlantic, group = 1), 
  #             method = "loess", color = "red", linewidth = 1.5, se = FALSE) +
  geom_hline(yintercept = 1, linetype = "dashed", color = "red", size = 0.5) +  # Horizontal line at y = 1
  # Line for average sex ratios across age groups 
  labs(#title = "Sex Ratios by Age Group and Year in Atlantic", 
    x = "Age", 
    y = "Sex Ratio (Male/Female)") +
  scale_x_discrete(labels = age_labels) +
  theme_minimal() +
  scale_fill_viridis_d(name = "Period",
                       labels = period_labels)  # Color bars by year

atlantic_sex_ratio_bar_chart


# Saving sex-ratio bar charts


pdf("sex_ratio_bar_charts.pdf", width = 11, height = 8)  

print(QC_sex_ratio_bar_chart)
print(ON_sex_ratio_bar_chart)
print(BC_sex_ratio_bar_chart)
print(AB_sex_ratio_bar_chart)
print(SK_sex_ratio_bar_chart)
print(MB_sex_ratio_bar_chart)
print(atlantic_sex_ratio_bar_chart)

dev.off()



# Tables for rates by cohort ####

library(readxl)
library(tidyverse)
library(stringr)
library(writexl)
library(openxlsx)

file_paths <- c("Counts/data_for_descriptive/male_ON_descriptive.xlsx",
                "Counts/data_for_descriptive/male_BC_descriptive.xlsx",
                "Counts/data_for_descriptive/male_QC_descriptive.xlsx",
                "Counts/data_for_descriptive/male_AB_descriptive.xlsx",
                "Counts/data_for_descriptive/male_SK_descriptive.xlsx",
                "Counts/data_for_descriptive/male_MB_descriptive.xlsx",
                "Counts/data_for_descriptive/male_atlantic_descriptive.xlsx",
                "Counts/data_for_descriptive/female_ON_descriptive.xlsx",
                "Counts/data_for_descriptive/female_BC_descriptive.xlsx",
                "Counts/data_for_descriptive/female_QC_descriptive.xlsx",
                "Counts/data_for_descriptive/female_AB_descriptive.xlsx",
                "Counts/data_for_descriptive/female_SKMB_descriptive.xlsx",
                "Counts/data_for_descriptive/female_atlantic_descriptive.xlsx")

table1_list <- list()

reference_age <- data.frame(age.n = rep(c(17, 22, 27, 32, 37, 42, 47, 52, 57, 62, 67, 72, 77, 82), times = 14),
                            yr.n =rep(c(1952, 1957, 1962, 1967, 1972, 1977, 1982, 1987, 1992, 1997,
                                        2002, 2007, 2012, 2017), each = 14))

for (file_path in file_paths) {
  
  dt <- read_xlsx(file_path) %>% 
    mutate(rate = round(rate, digits = 2))
  
  dt <- full_join(dt, reference_age, by = c("age.n", "yr.n"))
  
  if (!"cohort" %in%  colnames(dt)) {
    dt <- dt %>% 
      mutate(cohort = yr.n - age.n)
  }
  
  
  # Ensure dt has the required columns: age.n, yr.n, tot_p
  dt$age.f <- factor(dt$age.n)
  dt$period.f <- factor(dt$yr.n)
  dt$pop <- dt$tot_p
  
  a.n <- length(levels(dt$age.f))  # number of age groups
  p.n <- length(levels(dt$period.f))  # number of periods
  c.n <- a.n + p.n - 1  # number of cohorts
  
  table1 <- data.frame(
    age.group = rep(c("15-19", "20-24", "25-29", "30-34", 
                      "35-39", "40-44", "45-49", "50-54",
                      "55-59", "60-64", "65-69", "70-74",
                      "75-79", "80-84"), each = 2),
    y1950 = rep(NA, a.n * 2),
    y1955 = rep(NA, a.n * 2),
    y1960 = rep(NA, a.n * 2),
    y1965 = rep(NA, a.n * 2),
    y1970 = rep(NA, a.n * 2),
    y1975 = rep(NA, a.n * 2),
    y1980 = rep(NA, a.n * 2),
    y1985 = rep(NA, a.n * 2),
    y1990 = rep(NA, a.n * 2),
    y1995 = rep(NA, a.n * 2),
    y2000 = rep(NA, a.n * 2),
    y2005 = rep(NA, a.n * 2),
    y2010 = rep(NA, a.n * 2),
    y2015 = rep(NA, a.n * 2))
  
  
  # Defining the sequences for odd and even indices (28 is for the number of age rows * 2, one row for rate and one for cohort category)
  odd <- seq(1, 27, 2)
  even <- seq(2, 28, 2)
  
  s <- seq(1, p.n)
  for (i in s) {
    indices <- seq(i, i + 14 * (a.n - 1), 14)
    table1[odd, i + 1] <- dt$rate[indices]
    table1[even, i + 1] <- dt$cohort[indices]
  }
  
  table1_list[[file_path]] <- table1
  
}

# Writing the data a single excel doc with multiple sheets
# Create a new Excel workbook
wb <- createWorkbook()

# Loop through the list of tables
for (item_name in names(table1_list)) {
  
  base_name <- str_remove(str_remove(item_name, "Counts/data_for_descriptive/"), "_descriptive.xlsx")
  
  save_selected_table <- table1_list[[item_name]]
  
  # Add a new sheet to the workbook for each table
  addWorksheet(wb, base_name)
  
  # Write the table to the corresponding sheet
  writeData(wb, sheet = base_name, save_selected_table)
}

# Save the workbook to a file
saveWorkbook(wb, file = "Tables/Descriptive Stats/rates_for_all_provinces.xlsx", overwrite = TRUE)

# Predicted Suicide rates age vs APC ####

library(ggplot2)
library(readr)
library(readxl)
library(tidyverse)
library(ggthemes)

file_paths <- c(
  "final_data/female_atlantic.xlsx",
  "final_data/female_QC.xlsx",
  "final_data/female_ON.xlsx",
  "final_data/female_MB.xlsx",
  "final_data/female_SK.xlsx",
  "final_data/female_AB.xlsx",
  "final_data/female_BC.xlsx",
  "final_data/male_atlantic.xlsx",
  "final_data/male_QC.xlsx",
  "final_data/male_ON.xlsx",
  "final_data/male_MB.xlsx",
  "final_data/male_SK.xlsx",
  "final_data/male_AB.xlsx",
  "final_data/male_BC.xlsx"
)

# Store the graphs
predicted_plot <- list()

period.plot3 <- c("1950-1954", "1955-1959", "1960-1964", "1965-1969", 
                  "1970-1974", "1975-1979", "1980-1984", "1985-1989",
                  "1990-1994", "1995-1999", "2000-2004", "2005-2009", 
                  "2010-2014", "2015-2019")

age_age <- seq(17, 82, by = 5)

period.plot3 <- c("1950-1954", "1955-1959", "1960-1964", "1965-1969", 
                  "1970-1974", "1975-1979", "1980-1984", "1985-1989",
                  "1990-1994", "1995-1999", "2000-2004", "2005-2009", 
                  "2010-2014", "2015-2019")

# Create the Model vector
Model <- rep(c("Age main effects", "APC-I"), each = length(age_age) * length(period.plot3))

# Create the yr vector
yr <- rep(period.plot3, times = 14)

# Create the age_plot3 vector
age_plot3 <- rep(age_age, each = 14)

# Plot the results for predicted rates

for (file_path in file_paths) {
  
  dt <- read_xlsx(file_path)
  
  #Combine p1 and p3 into a single vector
  p1_plot3 <- dt$p1
  p3_plot3 <- dt$p3
  rate_plot3 <- c(p1_plot3, p3_plot3)
  
  # Combine all vectors into a dataframe
  dt.plot3 <- data.frame(rate_plot3 = rate_plot3, 
                         Model = Model, 
                         yr = yr, 
                         age_plot3 = age_plot3)
  
  # Adding the province to the title, from the file_path
  title <- str_remove(basename(file_path), "_final_data.xlsx") %>% 
    str_replace_all("_", " ") %>% 
    str_to_title()
  
  predicted_effects_plot <- ggplot(dt.plot3, aes(x = age_plot3, y = rate_plot3, group = Model, shape = Model, colour = Model)) +
    geom_line(aes(linetype = Model)) +
    geom_point() +
    facet_wrap(~yr, scales = "free_y") +
    scale_x_continuous(breaks = seq(17,82, by = 10),
                       labels = c("15-9", "25-9", 
                                  "35-9",  "45-9", 
                                  "55-9",  "65-9",
                                  "75-9")) +
    xlab("Age") + ylab("Predicted Rate") +
    theme(legend.position = c(0.9, 0.01), 
          legend.justification = c(1, 0),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          panel.background = element_rect(fill = "white"),
          panel.border = element_rect(colour = "black", fill = NA),
          strip.background = element_rect(colour = "black", fill = "grey85")
    ) +   
    scale_color_colorblind() +
    ggtitle(title)
  
  predicted_plot[[file_path]] <- predicted_effects_plot
  
}

print(predicted_plot)

pdf("predicted_effects_by_period.pdf", width = 10, height = 6)

print(predicted_plot)

dev.off()


# Predicted rates AP vs APC, by COHORT ####

library(ggplot2)
library(readr)
library(readxl)
library(tidyverse)
library(ggthemes)
library(ggrepel)

file_paths_m <- c(
  "final_data/male_atlantic.xlsx",
  "final_data/male_QC.xlsx",
  "final_data/male_ON.xlsx",
  "final_data/male_MB.xlsx",
  "final_data/male_SK.xlsx",
  "final_data/male_AB.xlsx",
  "final_data/male_BC.xlsx"
)

file_paths_f <- c(
  "final_data/female_atlantic.xlsx",
  "final_data/female_QC.xlsx",
  "final_data/female_ON.xlsx",
  "final_data/female_MB.xlsx",
  "final_data/female_SK.xlsx",
  "final_data/female_AB.xlsx",
  "final_data/female_BC.xlsx"
)

# Store the graphs
predicted_plot_cohort_m <- list()
predicted_plot_cohort_f <- list()

age_age <- seq(17, 82, by = 5)

period.plot_ap <- c("1950-1954", "1955-1959", "1960-1964", "1965-1969", 
                    "1970-1974", "1975-1979", "1980-1984", "1985-1989",
                    "1990-1994", "1995-1999", "2000-2004", "2005-2009", 
                    "2010-2014", "2015-2019")

# Create the Model vector
Model <- rep(c("AP", "APC-I"), each = length(age_age) * length(period.plot_ap)) 

# Create the yr vector
yr <- rep(period.plot_ap, times = 14)

# Create the age_plot3 vector
age_period_plot <- rep(age_age, each = 14)

# MALE - Plot the results for predicted rates 

for (file_path_m in file_paths_m) {
  
  final_data_m <- read_xlsx(file_path_m)
  
  #Combine p1 and p3 into a single vector
  p2_plot_ap <- final_data_m$p2
  p3_plot_ap <- final_data_m$p3
  rate_plot_ap <- c(p2_plot_ap, p3_plot_ap)
  
  # Combine all vectors into a dataframe
  dt.plot_ap <- data.frame(rate_plot_ap = rate_plot_ap, 
                           Model = Model, 
                           yr = yr, 
                           age_period_plot = age_period_plot)
  
  # Adding the province to the title, from the file_path
  title <- str_remove(basename(file_path_m), "_final_data.xlsx") %>% 
    str_replace_all("_", " ") %>% 
    str_to_title()
  
  dt.plot_1965 <- dt.plot_ap %>%
    mutate(
      yr.n = rep(seq(1952, 2017, by = 5), length.out = n()),
      cohort = yr.n - age_period_plot,
      cohort_start = cohort,
      cohort_end = cohort + 4,
      cohort_range = paste0(cohort_start, "-", cohort_end) 
    ) %>%
    filter(cohort %in% c(1940:1980)) %>%
    group_by(age_period_plot, cohort) %>%
    mutate(
      ap_rate = ifelse(Model == "AP", rate_plot_ap, NA),
      apc_rate = ifelse(Model == "APC-I", rate_plot_ap, NA),
      ap_rate = max(ap_rate, na.rm = TRUE), # Ensure `ap_rate` is filled
      percent_diff = ifelse(Model == "APC-I", (rate_plot_ap - ap_rate) / ap_rate * 100, NA)
    ) %>%
    ungroup()
  
  # Plot
  predicted_effects_plot_cohort_m <- ggplot(
    dt.plot_1965, 
    aes(x = age_period_plot, y = rate_plot_ap, group = Model, shape = Model, colour = Model)) +
    geom_line(aes(linetype = Model)) +
    geom_point() +
    geom_text_repel(
      data = dt.plot_1965 %>% filter(Model == "APC-I"),
      aes(label = sprintf("%.1f%%", percent_diff)),  
      colour = "black",   # Set font color to black
      size = 3,           # Adjust text size
      box.padding = 0.2,  # Space around the text
      point.padding = 0.3, # Space between points and text
      max.overlaps = Inf  # Avoid overlap
    ) +
    facet_wrap(~cohort_range, scales = "free_y") +
    scale_x_continuous(
      breaks = seq(17, 77, by = 5),
      labels = c(
        "15-19", "20-24", "25-29", 
        "30-34", "35-39", "40-44", 
        "45-49", "50-54", "55-59", 
        "60-64", "65-69", "70-74",
        "75-79"
      )
    ) +
    xlab("Age") + ylab("Predicted Rate") +
    theme(
      # legend.position = c(0.9, 0.05),
      # legend.justification = c(1, 0),
      panel.grid.major = element_blank(),
      panel.grid.minor = element_blank(),
      panel.background = element_rect(fill = "white"),
      panel.border = element_rect(colour = "black", fill = NA),
      strip.background = element_rect(colour = "black", fill = "grey85"),
      axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1)
    ) +
    scale_color_colorblind() +
    ggtitle(title)
  
  predicted_plot_cohort_m[[file_path_m]] <- predicted_effects_plot_cohort_m
  
}


pdf("predicted_effects_by_cohort_male.pdf", width = 10, height = 6)

print(predicted_plot_cohort_m)

dev.off()


### ________________________________ ###

# FEMALE predicted rates

for (file_path_f in file_paths_f) {
  
  final_data_f <- read_xlsx(file_path_f)
  
  #Combine p1 and p3 into a single vector
  p2_plot_ap <- final_data_f$p2
  p3_plot_ap <- final_data_f$p3
  rate_plot_ap <- c(p2_plot_ap, p3_plot_ap)
  
  # Combine all vectors into a dataframe
  dt.plot_ap <- data.frame(rate_plot_ap = rate_plot_ap, 
                           Model = Model, 
                           yr = yr, 
                           age_period_plot = age_period_plot)
  
  # Adding the province to the title, from the file_path
  title <- str_remove(basename(file_path_f), "_final_data.xlsx") %>% 
    str_replace_all("_", " ") %>% 
    str_to_title()
  
  dt.plot_1965 <- dt.plot_ap %>%
    mutate(
      yr.n = rep(seq(1952, 2017, by = 5), length.out = n()),
      cohort = yr.n - age_period_plot,
      cohort_start = cohort,
      cohort_end = cohort + 4,
      cohort_range = paste0(cohort_start, "-", cohort_end)
    ) %>%
    filter(cohort %in% c(1955:1995)) %>%
    group_by(age_period_plot, cohort) %>%
    mutate(
      ap_rate = ifelse(Model == "AP", rate_plot_ap, NA),
      apc_rate = ifelse(Model == "APC-I", rate_plot_ap, NA),
      ap_rate = max(ap_rate, na.rm = TRUE), # Ensure `ap_rate` is filled
      percent_diff = ifelse(Model == "APC-I", (rate_plot_ap - ap_rate) / ap_rate * 100, NA)
    ) %>%
    ungroup()
  
  # Plot
  predicted_effects_plot_cohort_f <- ggplot(
    dt.plot_1965, 
    aes(x = age_period_plot, y = rate_plot_ap, group = Model, shape = Model, colour = Model)) +
    geom_line(aes(linetype = Model)) +
    geom_point() +
    geom_text_repel(
      data = dt.plot_1965 %>% filter(Model == "APC-I"),
      aes(label = sprintf("%.1f%%", percent_diff)),  
      colour = "black",   # Set font color to black
      size = 3,           # Adjust text size
      box.padding = 0.2,  # Space around the text
      point.padding = 0.3, # Space between points and text
      max.overlaps = Inf  # Avoid overlap
    ) +
    facet_wrap(~cohort_range, scales = "free_y") +
    scale_x_continuous(
      breaks = seq(17, 62, by = 5),
      labels = c(
        "15-19", "20-24", "25-9", 
        "30-34", "35-39", "40-44", 
        "45-9", "50-54", "55-59", "60-64"
      )
    ) +
    xlab("Age") + ylab("Predicted Rate") +
    theme(
      # legend.position = c(0.9, 0.05),
      # legend.justification = c(1, 0),
      panel.grid.major = element_blank(),
      panel.grid.minor = element_blank(),
      panel.background = element_rect(fill = "white"),
      panel.border = element_rect(colour = "black", fill = NA),
      strip.background = element_rect(colour = "black", fill = "grey85"),
      axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1)
    ) +
    scale_color_colorblind() +
    ggtitle(title)
  
  predicted_plot_cohort_f[[file_path_f]] <- predicted_effects_plot_cohort_f
  
}

# Save both M and F on the same pdf
print(predicted_plot_cohort_f)

pdf("predicted_effects_by_cohort_combined_mf.pdf", width = 10, height = 6)

print(predicted_plot_cohort_m)
print(predicted_plot_cohort_f)

dev.off()

# Heatmaps for cohort interaction terms ####

library(tidyverse)
library(readxl)
library(viridis) #colours for graphs, specifically colour blind pallets


file_paths <- c("MASTER/Sortie_cohort Coefficient/male_QC_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/male_ON_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/male_BC_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/male_SK_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/male_MB_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/male_AB_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/male_atlantic_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/female_QC_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/female_ON_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/female_BC_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/female_SK_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/female_MB_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/female_AB_inter_coh_coef_by_age_period.xlsx",
                "MASTER/Sortie_cohort Coefficient/female_atlantic_inter_coh_coef_by_age_period.xlsx")

age_ranges <- c("15-19", "20-24", "25-29", "30-34", "35-39", 
                "40-44", "45-49", "50-54", "55-59", "60-64", 
                "65-69", "70-74", "75-79", "80-84")

periods <- c("1950-1954", "1955-1959", "1960-1964", "1965-1969", 
             "1970-1974", "1975-1979", "1980-1984", "1985-1989", 
             "1990-1994", "1995-1999", "2000-2004", "2005-2009", 
             "2010-2014", "2015-2019")

heatmap_plot <- list()

for (file_path in file_paths) {
  
  cco_df <- read_xlsx(file_path)
  
  pivot_cc0 <- cco_df %>% 
    pivot_longer(
      cols = starts_with("c"),
      names_to = "cohort_dev",
      values_to = "coefficient"
    ) %>% 
    select(age.group, coefficient) %>% 
    dplyr::mutate(periods = rep(periods, times = 14))
  
  # Detecting outliers
  Q1 <- quantile(pivot_cc0$coefficient, 0.25, na.rm = TRUE)
  Q3 <- quantile(pivot_cc0$coefficient, 0.75, na.rm = TRUE)
  IQR <- Q3 - Q1
  lower_bounds <- Q1 - 1.5 * IQR
  upper_bounds <- Q3 + 1.5 * IQR
  has_outliers <- any(pivot_cc0$coefficient < lower_bounds | pivot_cc0$coefficient > upper_bounds, na.rm = TRUE)
  
  # Remove outliers only if they exist
  if (has_outliers) {
    pivot_cc0 <- pivot_cc0 %>%
      dplyr::filter(coefficient >= lower_bounds & coefficient <= upper_bounds)
  }
  
  # Complete data frame to ensure all periods and age groups are represented
  complete_df <- data.frame(periods = rep(periods, times = 14),
                            age.group = rep(age_ranges, each = 14))
  
  # Merge and retain NA values where outliers were removed
  merged_cco <- merge(complete_df, pivot_cc0, by = c("periods", "age.group"), all.x = TRUE)
  
  # Extract province name from file path for title
  title <- str_remove(basename(file_path), "_inter_coh_coef_by_age_period.xlsx") %>% 
    str_replace_all("_", " ") %>% 
    str_to_title()
  
  # Generate color-blind-friendly heatmap
  colour_blind_heatmap <- ggplot(merged_cco, aes(x=periods, y=age.group, fill=coefficient)) +
    geom_tile(color="white") +
    scale_fill_viridis(option="plasma", name="Coefficient") +  # "plasma" is one of the colorblind-friendly options
    theme_minimal() + 
    theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1)) +
    labs(title = paste0(title, " IQR"),
         x = "Period",
         y = "Age")
  
  heatmap_plot[[title]] <- colour_blind_heatmap
  
}

# Heatmap with all values

heatmap_2 <- list()

for (file_path in file_paths) {
  
  cco_df <- read_xlsx(file_path)
  
  pivot_cc0 <- cco_df %>% 
    pivot_longer(
      cols = starts_with("c"),
      names_to = "cohort_dev",
      values_to = "coefficient"
    ) %>% 
    select(age.group, coefficient) %>% 
    dplyr::mutate(periods = rep(periods, times = 14))
  
  # Extract province name from file path for title
  title <- str_remove(basename(file_path), "_inter_coh_coef_by_age_period.xlsx") %>% 
    str_replace_all("_", " ") %>% 
    str_to_title()
  
  # Generate color-blind-friendly heatmap
  colour_blind_heatmap <- ggplot(pivot_cc0, aes(x=periods, y=age.group, fill=coefficient)) +
    geom_tile(color="white") +
    scale_fill_viridis(option="plasma", name="Coefficient") +  # "plasma" is one of the colorblind-friendly options
    theme_minimal() + 
    theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1)) +
    labs(title = title,
         x = "Period",
         y = "Age")
  
  heatmap_2[[title]] <- colour_blind_heatmap
  
}

pdf("cohort_heatmap.pdf", width = 10, height = 6)

print(heatmap_plot)

print(heatmap_2)

dev.off() 


# Age, period, and cohort coefficients plot ####

library(ggplot2)
library(tidyverse)
library(patchwork)
library(readxl)
library(stringr)

file_paths <- c(
  "final_data/female_atlantic.xlsx",
  "final_data/female_QC.xlsx",
  "final_data/female_ON.xlsx",
  "final_data/female_MB.xlsx",
  "final_data/female_SK.xlsx",
  "final_data/female_AB.xlsx",
  "final_data/female_BC.xlsx",
  "final_data/male_atlantic.xlsx",
  "final_data/male_QC.xlsx",
  "final_data/male_ON.xlsx",
  "final_data/male_MB.xlsx",
  "final_data/male_SK.xlsx",
  "final_data/male_AB.xlsx",
  "final_data/male_BC.xlsx"
)

file_path_cohort <- c(
  "inter_cohort_dev/female_atlantic_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/female_QC_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/female_ON_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/female_MB_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/female_SK_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/female_AB_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/female_BC_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/male_atlantic_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/male_QC_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/male_ON_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/male_MB_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/male_SK_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/male_AB_inter_cohort_deviation.xlsx",
  "inter_cohort_dev/male_BC_inter_cohort_deviation.xlsx"
)

# Initialize lists for storing plots
final_plot_list <- list()
age_plot_list <- list()
period_plot_list <- list()
cohort_plot_list <- list()

# Loop 
for (i in seq_along(file_paths)) {
  file_path <- file_paths[i]
  cohort_path <- file_path_cohort[i]
  
  # Extract base names
  base_name <- str_remove(basename(file_path), ".xlsx")
  
  # Read the data
  data <- read_xlsx(file_path)
  
  # Cleaning the cohort dev data and creating the x-axis labels
  cohort_data <- read_xlsx(cohort_path) %>%
    select(cohort, deviation) %>%
    mutate(
      cohort_up = cohort + 4,
      cohort_down = cohort,
      cohort = paste0(cohort_down, "-", cohort_up)
    )
  
  # Create age plot
  age_plot <- ggplot(data, aes(x = age.n, y = aco)) + 
    geom_point() +
    geom_line() +
    geom_hline(yintercept = 0) +
    scale_x_continuous(
      breaks = seq(17, 82, by = 5),
      labels = c("15-19", "20-24", "25-29", "30-34", "35-39", "40-44",
                 "45-49", "50-54", "55-59", "60-64", "65-69", "70-74",
                 "75-79", "80-84")
    ) +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1)) +
    xlab("Age") + ylab("Age Effect") +
    ggtitle("Estimated Age Effects")
  
  # Create period plot
  period_plot <- ggplot(data, aes(x = yr.n, y = pco)) +
    geom_point() +
    geom_line() +
    geom_hline(yintercept = 0) +
    scale_x_continuous(
      breaks = seq(1952, 2017, 5),
      labels = c("1950-1954", "1955-1959", "1960-1964", "1965-1969",
                 "1970-1974", "1975-1979", "1980-1984", "1985-1989",
                 "1990-1994", "1995-1999", "2000-2004", "2005-2009",
                 "2010-2014", "2015-2019")
    ) +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1)) +
    xlab("Period") + ylab("Period Effect") +
    ggtitle("Estimated Period Effects")
  
  # Create cohort plot
  cohort_plot <- ggplot(cohort_data, aes(x = cohort, y = deviation, group = 1)) +
    geom_point() +
    geom_line() +
    geom_hline(yintercept = 0) +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1)) +
    xlab("Cohort") + ylab("Inter-Cohort Deviation") +
    ggtitle("Estimated Inter-Cohort Deviation Effects")
  
  # Combine plots 
  layout <- (age_plot | period_plot) 
  / cohort_plot
  
  # Store plots
  final_plot_list[[base_name]] <- layout
  age_plot_list[[base_name]] <- age_plot
  period_plot_list[[base_name]] <- period_plot
  cohort_plot_list[[base_name]] <- cohort_plot
  
}

# Check plots
final_plot_list

# Save plots


# Directory to save the plots
output_dir <- "Figures/Combined_APC_effects/"

for (base_name in names(final_plot_list)) {
  
  plot <- final_plot_list[[base_name]]
  
  # Save as PNG (high resolution)
  ggsave(
    filename = file.path(output_dir, paste0(base_name, ".png")),
    plot = plot,
    width = 10, height = 8, dpi = 300
  )
  
  # Optionally, save as PDF for vector graphics
  ggsave(
    filename = file.path(output_dir, paste0(base_name, ".pdf")),
    plot = plot,
    width = 10, height = 8
  )
}





