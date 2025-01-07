# In case you are finding it difficult to understand this code, here is my email: kamarasekou798@gmail.com

# Data Preparation --------------------------------------------------------

# Load the necessary library
library(readxl)

# Read the dataset for Jubilee Holdings Ltd
JUB <- read_excel("C:/Jubilee_Insurance_Markov/JUB.xlsx")

# Function to arrange data in ascending order (assumes data is initially in descending order by year)
order_data <- function(stock_data) {
  ordered_data <- c()
  pointer <- length(stock_data)
  while (length(stock_data) > length(ordered_data)) {
    ordered_data <- c(ordered_data, stock_data[pointer])
    pointer <- pointer - 1
  }
  return(ordered_data)
}

# Apply ordering to Jubilee Holdings share price closing data
jub_data <- order_data(JUB$Closing)

# Transition Analysis -----------------------------------------------------

# Function to generate state sequence based on the first difference of the data
generate_state_sequence <- function(transition_data) {
  pointer <- 1
  stateVec <- c()
  for (i in transition_data) {
    pointer <- pointer + 1
    if (length(transition_data) != (pointer - 1)) {
      if (i > transition_data[pointer]) {
        stateVec <- c(stateVec, "Down")
      } else if (i < transition_data[pointer]) {
        stateVec <- c(stateVec, "Up")
      } else {
        stateVec <- c(stateVec, "No-change")
      }
    }
  }
  return(stateVec)
}

# Generate the state sequence
tran_sec <- generate_state_sequence(jub_data)
states <- c("Up", "Down", "No-change")


# Generate transition frequency,transition matrix, Stationary Distribution, initial distribution and convergence table
transition_matrix <- function(states, transition_data, conv.num = 10) {
  # Initialize storage for transition counts and probabilities
  transition_list <- vector("list", length = length(states))
  transition_list_2 <- vector("list", length = length(states))
  total_transitions <- numeric(length = length(states))
  names(transition_list) <- states
  names(transition_list_2) <- states
  
  # Compute transition counts and probabilities
  for (name in states) {
    pointer <- 1
    next_states <- c()
    for (j in transition_data) {
      pointer <- pointer + 1
      if (name == j & length(transition_data) != pointer - 1) {
        next_states <- c(next_states, transition_data[pointer])
      }
    }
    transition_list[[name]] <- t(as.matrix(prop.table(table(factor(next_states, levels = states)))))
    transition_list_2[[name]] <- t(as.matrix(table(factor(next_states, levels = states))))
  }
  
  # Create transition matrix and count matrix
  transition_matrix <- matrix(0, ncol = length(states), byrow = TRUE)
  transition_count <- matrix(0, ncol = length(states), byrow = TRUE)
  
  pointer <- 1
  for (name in states) {
    transition_matrix <- rbind(transition_matrix, transition_list[[name]])
    transition_count <- rbind(transition_count, transition_list_2[[name]])
    total_transitions[pointer] <- sum(transition_list_2[[name]])
    pointer <- pointer + 1
  }
  transition_matrix <- transition_matrix[-1,]
  rownames(transition_matrix) <- states
  colnames(transition_matrix) <- states
  
  transition_count <- transition_count[-1,]
  transition_count <- cbind(transition_count, total_transitions)
  rownames(transition_count) <- states
  colnames(transition_count) <- c(states, "Total_transition")
  
  # Check transition matrix validity
  for (name in states) {
    if (sum(transition_matrix[name,]) != 1) {
      stop("Transition matrix rows not adding up to one")
    }
  }
  
  # Compute long-term distribution
  longterm_dist <- function(tran_matrix) {
    identity_mat <- diag(ncol(tran_matrix))
    add_row <- rep(1, ncol(tran_matrix))
    coef_matrix <- rbind((t(tran_matrix) - identity_mat)[-1,], add_row)
    out_vec <- c(rep(0, ncol(tran_matrix) - 1), 1)
    long_dist <- solve(coef_matrix, out_vec)
    return(long_dist)
  }
  stationary_dist <- longterm_dist(transition_matrix)
  
  # Compute initial distribution
  initial_dist <- t(as.matrix(prop.table(table(factor(transition_data, levels = states)))))
  
  # Convergence table using matrix powers
  library(expm)
  convergence_table <- matrix(0, ncol = length(states), byrow = TRUE)
  steps <- c()
  for (i in 1:conv.num) {
    conv_mat <- initial_dist %*% (transition_matrix %^% i)
    convergence_table <- rbind(convergence_table, conv_mat)
    steps[i] <- i
  }
  convergence_table <- convergence_table[-1,]
  convergence_table <- cbind(convergence_table, steps)
  colnames(convergence_table) <- c(states, "Day(s)_from_now")
  
  # Round results for better readability
  transition_count <- round(transition_count, 3)
  transition_matrix <- round(transition_matrix, 3)
  stationary_dist <- round(stationary_dist, 3)
  initial_dist <- round(initial_dist, 3)
  convergence_table <- round(convergence_table, 3)
  
  return(list(transition_count, transition_matrix, stationary_dist, initial_dist, convergence_table))
}

# Apply the transition matrix function
tran_matrix <- transition_matrix(states, tran_sec)


# Generate simulation for the markov process
simulate_markov <- function(tran_matrix, init_dist, states, n_steps = 1000) {
  # Initialize the process with a starting state sampled from the initial distribution
  cumulative_vec <- sample(states, 1, prob = init_dist)
  
  # Iterate over the number of steps to simulate the process
  for (i in 1:n_steps) {
    for (name in states) {
      # Determine the next state based on the transition matrix probabilities
      if (cumulative_vec[i] == name) {
        cumulative_vec <- c(cumulative_vec, sample(states, 1, prob = tran_matrix[name, ]))
      }
    }
  }
  
  # Map state names to numerical values for plotting
  cumulative_vec[cumulative_vec == "Up"] <- 1
  cumulative_vec[cumulative_vec == "Down"] <- -1
  cumulative_vec[cumulative_vec == "No-change"] <- 0
  cumulative_vec <- as.integer(cumulative_vec)
  
  # Compute the cumulative sum of the numerical values
  cumulative_vec <- cumsum(as.integer(cumulative_vec))
  
  # Plot the simulation results
  plot(seq(0, n_steps, 1), cumulative_vec, type = "s", col = "brown",
       xlab = "Steps", ylab = "Positions")
}

# Example usage of the function
simulate_markov(tran_matrix[[2]], tran_matrix[[4]], states)


# Generate the table of observed and expected transitions for the triplet test
transition_fit_dataframe <- function(states, transition_data, tran_matrix) {
  # Create a list of transition names (e.g., "UpUp", "UpDown")
  transition_names <- numeric(length = length(states)**2)
  pointer <- 1
  for (name_one in states) {
    for (name_two in states) {
      transition_names[pointer] <- paste0(name_one, name_two)
      pointer <- pointer + 1
    }
  }
  
  # Initialize storage for transition data
  total_transitions <- numeric(length = length(states)**2)
  transition_list <- vector("list", length = length(states)**2)
  names(transition_list) <- transition_names
  
  # Analyze observed transitions in the data
  for (name_one in states) {
    for (name_two in states) {
      pointer <- 2
      next_states <- c()
      for (j in 2:length(transition_data)) {
        pointer <- pointer + 1
        if (transition_data[j - 1] == name_one && transition_data[j] == name_two &&
            length(transition_data) != (pointer - 1)) {
          next_states <- c(next_states, transition_data[pointer])
        }
      }
      transition_list[[paste0(name_one, name_two)]] <- 
        t(as.matrix(table(factor(next_states, levels = states))))
    }
  }
  
  # Compile observed transitions into a matrix
  transition_count <- matrix(0, ncol = length(states), byrow = TRUE)
  colnames(transition_count) <- states
  
  pointer <- 1
  for (name in transition_names) {
    transition_count <- rbind(transition_count, transition_list[[name]])
    total_transitions[pointer] <- sum(transition_list[[name]])
    pointer <- pointer + 1
  }
  transition_count <- transition_count[-1, ]
  
  # Calculate expected transition frequencies
  expected_freq <- c()
  pointer <- 1
  for (i in 1:(length(states)**2)) {
    if (pointer > length(states)) {
      pointer <- 1
    }
    expected_freq <- c(expected_freq, total_transitions[i] * tran_matrix[states[pointer], ])
    pointer <- pointer + 1
  }
  expected_matrix <- matrix(expected_freq, ncol = length(states), byrow = TRUE)
  
  # Append expected frequencies and totals to the transition matrix
  expected_names <- sapply(states, function(state) paste0("Expected-", state))
  transition_count <- cbind(transition_count, expected_matrix, total_transitions)
  rownames(transition_count) <- transition_names
  colnames(transition_count) <- c(states, expected_names, "Total_transition")
  
  return(transition_count)
}

# Example usage
triple <- transition_fit_dataframe(states, tran_sec, tran_matrix[[2]])


# Chi-Square Test for Markov Model Fit 
chisq_test <- function(states, triplet_data, df = 24, sig_level = 0.05) {
  library(glue)
  
  # Extract observed and expected frequencies
  test_observed <- c()
  test_expected <- c()
  for (j in 1:length(states)) {
    test_observed <- c(test_observed, triplet_data[, j])
    test_expected <- c(test_expected, triplet_data[, j + length(states)])
  }
  
  # Calculate chi-square statistic
  chisq_sequence <- ((test_observed - test_expected)**2) / test_expected
  chisq_sequence[is.nan(chisq_sequence) | !is.finite(chisq_sequence)] <- 0
  
  chisq_statistics <- round(sum(chisq_sequence), 3)
  chisq_critical <- round(qchisq(sig_level, df, lower.tail = FALSE), 3)
  chisq_pvalue <- round(pchisq(chisq_statistics, df, lower.tail = FALSE), 3)
  
  # Interpret test result
  if (chisq_critical < chisq_statistics) {
    glue("Null hypothesis: The data fit the Markov model \n\n",
         "Alternative hypothesis: The data do not fit the Markov model \n\n",
         "P-value: {chisq_pvalue}\n",
         "Chi-square statistic: {chisq_statistics}\n",
         "Chi-square critical value: {chisq_critical}\n\n",
         "Test result: We fail to accept the null hypothesis")
  } else {
    glue("Null hypothesis: The data fit the Markov model \n\n",
         "Alternative hypothesis: The data do not fit the Markov model \n\n",
         "P-value: {chisq_pvalue}\n",
         "Chi-square statistic: {chisq_statistics}\n",
         "Chi-square critical value: {chisq_critical}\n\n",
         "Test result: We fail to reject the null hypothesis")
  }
}

# Example usage
chisq_test(states, triple)

# Increment analysis ------------------------------------------------------


# Function to remove outliers from share price increments
outlier_removal <- function(stock_data, plot = FALSE, lower_bond = -40, upper_bond = 40) {
  # Calculate the increments (differences between consecutive stock prices)
  increments <- diff(stock_data)
  
  # Plot the increments with outliers if plot is TRUE
  if (plot == TRUE) {
    par(mfrow = c(2, 1), mar = c(2, 2, 2, 2))
    plot(increments, type = "l", col = "brown", main = "Increment with outliers")
  }
  
  # Remove outliers based on specified bounds
  increments <- increments[increments <= upper_bond & increments >= lower_bond]
  
  # Plot the increments without outliers if plot is TRUE
  if (plot == TRUE) {
    plot(increments, type = "l", col = "brown", main = "Increment without outliers - stationary")
  } 
  
  return(increments)
}

# Apply the function and visualize the results
incre_out <- outlier_removal(jub_data, plot = TRUE)


# Function to test stationarity or autocorrelation in stock price increments
auto_stationary_test <- function(stock_increment, t_test, lak) {
  library(tseries)
  
  # Perform stationarity test (ADF test) if t_test is "s" or "stationarity"
  if (t_test == "s" || t_test == "stationarity") {
    data_test <- adf.test(ts(stock_increment), alternative = "stationary", k = lak)
    
    # Perform autocorrelation test (Ljung-Box test) if t_test is "a" or "auto"
  } else if (t_test == "a" || t_test == "auto") {
    data_test <- Box.test(stock_increment, lag = lak, type = "Ljung-Box")
  }
  
  return(data_test)
}

# Conduct a stationarity test on the outlier-free increments
auto_stationary_test(incre_out, "s", 10)


# Function to conduct a Z-test for a given sample mean and assumed mean
z_test <- function(assumed_mean = 0, sample_mean, mean_sd, sig_level = 0.05) {
  library(glue)
  
  # Calculate Z-statistics, critical value, and p-value
  z_statistics <- round((sample_mean - assumed_mean) / mean_sd, 3)
  z_critical <- round(abs(qnorm(sig_level / 2)), 3)
  z_pvalue <- 2 * round(pnorm(abs(z_statistics), lower.tail = FALSE), 3)
  
  # Return the test result with detailed interpretation
  if (z_critical < z_statistics) {
    glue("Null hypothesis: The mean is {assumed_mean} \n\n",
         "Alternative hypothesis: The mean is not {assumed_mean} \n\n",
         "P-value: {z_pvalue}\n",
         "Z-statistics: {z_statistics}\n",
         "Z-critical value: +/-{z_critical}\n\n",
         "Test result: We fail to accept the null hypothesis")
  } else {
    glue("Null hypothesis: The mean is {assumed_mean} \n\n",
         "Alternative hypothesis: The mean is not {assumed_mean} \n\n",
         "P-value: {z_pvalue}\n",
         "Z-statistics: {z_statistics}\n",
         "Z-critical value: +/-{z_critical}\n\n",
         "Test result: We fail to reject the null hypothesis")
  }
}

# Conduct a Z-test with the given values
z_test(0, 0.072, sqrt(50.49243 / 3713), 0.05)


# Function to generate a histogram for the distribution of stock increments
generate_returns_dist <- function(stock_increments) {
  increments <- stock_increments
  
  # Calculate bin width using Freedman-Diaconis rule
  bin_width <- 2 * (IQR(increments) / (length(increments))^(1 / 3))
  number_bins <- (range(increments)[2] - range(increments)[1]) / bin_width
  number_bins <- ceiling(number_bins)
  
  # Reset graphics device to avoid overlapping plots
  dev.off()
  
  # Plot histogram of the increments
  hist(increments, breaks = number_bins, xlim = c(min(increments), max(increments)), xlab = "Ensembled values")
}

# Generate the histogram for outlier-free increments
generate_returns_dist(incre_out)

