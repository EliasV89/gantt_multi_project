import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.dates import date2num, DateFormatter
import matplotlib.dates as mdates
from datetime import datetime


# Define your custom color palette for phases
color_palette = {
    'Preparation': '#003f5c',
    'Pre-study': '#444e86',
    'Establishment': '#955196',
    'Implementation': '#dd5182',
    'Stock build-up': '#ff6e54',
    'Closure': '#ffa600',
}


# Read the project timeline data from Excel
def read_data_from_excel(file_name):
    df = pd.read_excel(file_name)
    return df

# Convert the data into a DataFrame
def prepare_data(df):
    data = []
    grouped = df.groupby('Project')
    for project, group in grouped:
        min_start = group['Start'].min()
        for _, row in group.iterrows():
            start_date = pd.to_datetime(row['Start'])
            end_date = pd.to_datetime(row['End'])
            # Add an invisible bar from the min_start to the phase start_date
            if start_date > min_start:
                data.insert(0, {
                    'Project': project,
                    'Phase': row['Phase'],
                    'Start': min_start,
                    'End': start_date,
                    'Gate': None,
                    'Visible': False
                })
            # Add visible bar
            data.insert(0, {
                'Project': project,
                'Phase': row['Phase'],
                'Start': start_date,
                'End': end_date,
                'Gate': row['Gate'],
                'Visible': True
            })
    df_prepared = pd.DataFrame(data)
    
    # Ensure consistent phase order
    all_phases = sorted(df_prepared['Phase'].unique())
    df_prepared['Phase'] = pd.Categorical(df_prepared['Phase'], categories=all_phases, ordered=True)
    
    return df_prepared

# Read data from the Excel file
df = read_data_from_excel('project_tasks.xlsx')

# Prepare the data for plotting
df = prepare_data(df)

# Define a toned-down cyberpunk-inspired color palette
cyberpunk_palette = ["#003f5c", "#7a5195", "#ef5675", "#D98880", "#7DCEA0", "#F39C73"]

# Calculate one month before today's date
one_month_before_today = pd.Timestamp.now() - pd.DateOffset(months=1)

# Update the plot_gantt function
def plot_gantt(df):
    projects = df['Project'].unique()
    num_projects = len(projects)
    
    plt.style.use('dark_background')
    fig, axes = plt.subplots(nrows=num_projects, ncols=1, sharex=True, figsize=(16, 9))

    fig.patch.set_facecolor('#2E2E2E')  # Warm gray background for the figure
    axes = [axes] if num_projects == 1 else axes

    # Get today's date and convert it to a number
    today = date2num(pd.to_datetime(datetime.today()))

    for i, (project, ax) in enumerate(zip(projects, axes)):
        project_data = df[df['Project'] == project]

        # Plot invisible bars from the min start date to the start of each phase
        for _, row in project_data.iterrows():
            if not row['Visible']:
                ax.barh(row['Phase'], date2num(row['End']) - date2num(row['Start']),
                        left=date2num(row['Start']), color='none', edgecolor='none', height=0.5, zorder=0)

        # Plot gridlines after setting zorder
        ax.grid(True, which='major', axis='x', linestyle='--', linewidth=0.5, color='white', zorder=1)
        ax.grid(False, axis='y')

        # Plot visible bars with the custom color palette
        for idx, row in project_data.iterrows():
            if row['Visible']:
                phase_color = color_palette.get(row['Phase'], '#FFFFFF')  # Default to white if phase not found
                ax.barh(row['Phase'], date2num(row['End']) - date2num(row['Start']),
                        left=date2num(row['Start']), color=phase_color, 
                        edgecolor='white', height=0.5, zorder=2)
                if date2num(row['End']) > today:
                    ax.text(date2num(row['End']), row['Phase'], f' | {row["Gate"]}', color='white', va='center', ha='left', fontsize=10, zorder=3)
                    ax.scatter(date2num(row['End']), row['Phase'], marker='o', s=12, color='white', zorder=4)

        # Add a vertical line indicating today's date
        ax.axvline(x=today, color='orange', linestyle='--', linewidth=1.5, zorder=0)

        # Set y-ticks and labels
        ax.set_yticks(df['Phase'].cat.categories)
        ax.set_yticklabels(df['Phase'].cat.categories)
        
        ax.set_title(project, color='white')
        ax.set_xlabel('')
        ax.set_ylabel('')
        ax.set_facecolor('#2E2E2E')  # Warm gray background for the axes
        ax.tick_params(axis='x', colors='white')
        ax.tick_params(axis='y', colors='white')

    # Set the x-axis limit to start from one month before today
    plt.xlim(date2num(one_month_before_today), None)
    # Set x-axis major ticks to the first day of each quarter and grid lines
    ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=(1, 4, 7, 10)))
    ax.xaxis.set_major_formatter(DateFormatter("%Y-%m-%d"))

    fig.autofmt_xdate()  # Rotate date labels for better readability
    
    plt.xlabel('Timeline', color='white')

    # Adjust subplot spacing to reduce vertical margins
    plt.subplots_adjust(hspace=0.25)  # Decrease the space between subplots
    
    plt.tight_layout()

    # Save the plot with a dynamic filename that includes the current date
    todays_date = datetime.today().strftime('%Y-%m-%d')
    plt.savefig(f'output/projects_timeline_{todays_date}.png', dpi=150)

# Plot the Gantt chart
plot_gantt(df)
