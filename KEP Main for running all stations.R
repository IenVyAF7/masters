# Install and load required packages
install.packages(c("ggplot2", "openxlsx"))
library(ggplot2)
library(openxlsx)

# Read the stations data from Excel
stations_data <- read.xlsx("Frame.xlsx")

# Parameters
num_simulations <- 10000
months <- 12
years <- 20  # Number of years for simulation

# Monthly solar radiation data (replace with your actual data)
monthly_radiation <- c(43.09,61.04,91.14,115.5, 138.57, 133.5,140.74,133.61,99.3,78.43,52.5,36.89)

# Define probability distribution for solar panel efficiency
mean_efficiency <- 0.20  # Example mean efficiency
sd_efficiency <- 0.02  # Example standard deviation

# Initialize a data frame to store results for each station
results_df <- data.frame(Station_name = character(), Mean_Yearly_Energy = numeric(), stringsAsFactors = FALSE)

# Loop over each station
for(station_index in 1:nrow(stations_data)) {
  station_name <- stations_data$Station_name[station_index]
  solar_panel_area <- stations_data$Frame_area[station_index]
  
  # Simulate solar panel efficiency for each simulation
  simulated_efficiency <- rnorm(num_simulations * months * years, mean = mean_efficiency, sd = sd_efficiency)
  
  # Calculate monthly energy production for each simulation
  simulated_energy <- matrix(NA, nrow = num_simulations, ncol = months * years)
  for (i in 1:num_simulations) {
    simulated_energy[i, ] <- monthly_radiation * simulated_efficiency[((i - 1) * months * years + 1):(i * months * years)]
  }
  simulated_energy <- simulated_energy * solar_panel_area
  
  # Calculate yearly energy production
  yearly_energy <- rowSums(matrix(simulated_energy, ncol = months))
  
  # Calculate mean yearly energy production for this station
  mean_yearly_energy <- mean(yearly_energy)
  
  # Store the results for this station in the results data frame
  results_df <- rbind(results_df, data.frame(Station_name = station_name, Mean_Yearly_Energy = mean_yearly_energy))
}

# Write the results to an Excel file
write.xlsx(results_df, "mean_yearly_energy_results.xlsx")
