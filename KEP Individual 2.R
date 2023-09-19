library(openxlsx)

# Load the data from the Excel file
stations_data <- read.xlsx("Data.xlsx")

# Cost-benefit analysis function
perform_cba <- function(station) {
  time_period <- 20
  discount_rate <- 0.06
  electricity_price_per_kWh <- 0.30
  
  costs <- numeric(time_period)
  benefits <- numeric(time_period)
  
  for (year in 1:time_period) {
    costs[year] <- as.numeric(station["Maintenace"])  # Using 'Maintenace' column for annual maintenance
    benefits[year] <- as.numeric(station["Mean_Yearly_Energy"]) * electricity_price_per_kWh  # Using 'Mean_Yearly_Energy' column for annual energy production
  }
  
  costs[1] <- costs[1] + as.numeric(station["Installation.Cost"])  # Sum of equipment, installation, and permitting fees
  
  npv_costs <- sum(costs / (1 + discount_rate)^(1:time_period))
  npv_benefits <- sum(benefits / (1 + discount_rate)^(1:time_period))
  
  bcr <- npv_benefits / npv_costs
  
  # Calculate payback period
  cumulative_net_benefits <- cumsum(benefits - costs)
  payback_period <- which(cumulative_net_benefits > 0)[1]
  
  if(is.null(payback_period)) {
    payback_period <- NA  # If the project never breaks even
  }
  
  return(c(npv_costs = npv_costs, npv_benefits = npv_benefits, bcr = bcr, payback_period = payback_period))
}

# Apply the function to each row (station) in the data frame
results <- t(apply(stations_data, 1, perform_cba))

# Convert results to a data frame
results_df <- as.data.frame(results)

# Save the results back to Excel
write.xlsx(results_df, "cba_results.xlsx")

# Summary and Visualization

library(dplyr) 

# Summary statistics
summary_stats <- results_df %>% summarise(
  avg_npv_costs = mean(npv_costs),
  avg_npv_benefits = mean(npv_benefits),
  avg_bcr = mean(bcr),
  median_npv_costs = median(npv_costs),
  median_npv_benefits = median(npv_benefits),
  median_bcr = median(bcr)
)

print(summary_stats)


# Read the results from Excel
results_df <- read.xlsx("cba_results.xlsx")

# Histogram for NPV Costs
ggplot(results_df, aes(x = npv_costs)) +
  geom_histogram(binwidth = (max(results_df$npv_costs) - min(results_df$npv_costs)) / 30, fill = "blue", color = "black") +
  coord_cartesian(xlim = c(quantile(results_df$npv_costs, 0.01), quantile(results_df$npv_costs, 0.99))) +
  labs(title = "Histogram of NPV Costs", x = "NPV Costs", y = "Count") +
  theme_minimal()

# Histogram for NPV Benefits
ggplot(results_df, aes(x = npv_benefits)) +
  geom_histogram(binwidth = (max(results_df$npv_benefits) - min(results_df$npv_benefits)) / 30, fill = "green", color = "black") +
  coord_cartesian(xlim = c(quantile(results_df$npv_benefits, 0.01), quantile(results_df$npv_benefits, 0.99))) +
  labs(title = "Histogram of NPV Benefits", x = "NPV Benefits", y = "Count") +
  theme_minimal()

# Histogram for Benefit-Cost Ratio (BCR)
ggplot(results_df, aes(x = bcr)) +
  geom_histogram(binwidth = (max(results_df$bcr) - min(results_df$bcr)) / 30, fill = "red", color = "black") +
  coord_cartesian(xlim = c(quantile(results_df$bcr, 0.01), quantile(results_df$bcr, 0.99))) +
  labs(title = "Histogram of Benefit-Cost Ratio (BCR)", x = "BCR", y = "Count") +
  theme_minimal()

# Scatter plot for NPV Costs vs NPV Benefits
ggplot(results_df, aes(x = npv_costs, y = npv_benefits)) +
  geom_point() +
  labs(title = "Scatter plot of NPV Costs vs NPV Benefits", x = "NPV Costs", y = "NPV Benefits") +
  theme_minimal()
