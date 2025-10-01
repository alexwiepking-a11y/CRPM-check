import os
from datetime import datetime, timedelta
from collections import defaultdict

def create_actionable_dashboard(results, output_dir, timestamp):
    """Create a truly actionable dashboard focused on decisions and quick wins"""
    dashboard_file = os.path.join(output_dir, f"CRPM_Dashboard_{timestamp}.html")
    
    # Calculate metrics
    total_entries = results['total_entries']
    compliance_rate = results['compliance_rate']
    issues_count = len(results['true_deviations'])
    accepted_count = len(results['accepted_deviations'])
    
    # ACTIONABLE ANALYSIS - Find the patterns that matter
    
    # 1. QUICK WINS - Hotels with many issues (80/20 rule)
    hotel_issue_count = defaultdict(int)
    hotel_details = defaultdict(lambda: {'issues': [], 'countries': set(), 'priorities': defaultdict(int)})
    
    for dev in results['true_deviations']:
        hotel = dev['Hotel_Code']
        hotel_issue_count[hotel] += 1
        hotel_details[hotel]['issues'].append(dev['Deviation_Details'])
        hotel_details[hotel]['countries'].add(dev['Country'])
        hotel_details[hotel]['priorities'][dev['Priority']] += 1
    
    # Sort by issue count - these are your quick wins
    top_problem_hotels = sorted(hotel_issue_count.items(), key=lambda x: x[1], reverse=True)[:10]
    
    # Calculate impact: fixing top 5 hotels = what % improvement?
    top_5_issues = sum(count for _, count in top_problem_hotels[:5])
    impact_percentage = (top_5_issues / issues_count * 100) if issues_count > 0 else 0
    
    # 2. PATTERN DETECTION - What rules would fix the most issues?
    pattern_impact = defaultdict(lambda: {'count': 0, 'hotels': set(), 'countries': set(), 'examples': []})
    
    for dev in results['true_deviations']:
        details = dev['Deviation_Details']
        if 'VAT:' in details:
            # Extract VAT pattern
            vat_part = details.split('VAT: ')[1].split(' |')[0]
            if ' → ' in vat_part:
                current, standard = vat_part.replace("'", "").split(' → ')
                pattern_key = f"VAT_{current}_to_{standard}"
                pattern_impact[pattern_key]['count'] += 1
                pattern_impact[pattern_key]['hotels'].add(dev['Hotel_Code'])
                pattern_impact[pattern_key]['countries'].add(dev['Country'])
                if len(pattern_impact[pattern_key]['examples']) < 3:
                    pattern_impact[pattern_key]['examples'].append(f"{dev['Hotel_Code']} ({dev['Country']})")
        
        if 'Subaccount:' in details:
            sub_part = details.split('Subaccount: ')[1].split(' |')[0]
            if ' → ' in sub_part:
                current, standard = sub_part.replace("'", "").split(' → ')
                pattern_key = f"Subaccount_{current}_to_{standard}"
                pattern_impact[pattern_key]['count'] += 1
                pattern_impact[pattern_key]['hotels'].add(dev['Hotel_Code'])
                pattern_impact[pattern_key]['countries'].add(dev['Country'])
                if len(pattern_impact[pattern_key]['examples']) < 3:
                    pattern_impact[pattern_key]['examples'].append(f"{dev['Hotel_Code']} ({dev['Country']})")
    
    # Top patterns that would have biggest impact
    top_patterns = sorted(pattern_impact.items(), key=lambda x: x[1]['count'], reverse=True)[:5]
    
    # 3. EFFICIENCY ANALYSIS - How well are exception rules working?
    rule_efficiency = {
        'total_processed': total_entries,
        'auto_approved': accepted_count,
        'needs_review': issues_count,
        'efficiency_rate': (accepted_count / (accepted_count + issues_count) * 100) if (accepted_count + issues_count) > 0 else 100,
        'manual_work_remaining': issues_count
    }
    
    # 4. COUNTRY/REGION FOCUS - Which regions need attention?
    country_analysis = defaultdict(lambda: {'issues': 0, 'hotels': set(), 'types': defaultdict(int)})
    for dev in results['true_deviations']:
        country = dev['Country']
        country_analysis[country]['issues'] += 1
        country_analysis[country]['hotels'].add(dev['Hotel_Code'])
        if 'VAT:' in dev['Deviation_Details']:
            country_analysis[country]['types']['VAT'] += 1
        if 'Subaccount:' in dev['Deviation_Details']:
            country_analysis[country]['types']['Subaccount'] += 1
        if 'City Tax:' in dev['Deviation_Details']:
            country_analysis[country]['types']['City Tax'] += 1
    
    top_countries = sorted(country_analysis.items(), key=lambda x: x[1]['issues'], reverse=True)[:5]
    
    # Create the HTML with ACTIONABLE insights
    html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CRPM Action Dashboard - {timestamp}</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }}
        .dashboard {{
            background: white;
            border-radius: 15px;
            padding: 30px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            max-width: 1400px;
            margin: 0 auto;
        }}
        .header {{
            text-align: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 3px solid #667eea;
        }}
        .header h1 {{
            color: #333;
            margin: 0;
            font-size: 2.5rem;
        }}
        .subtitle {{
            color: #666;
            font-size: 1.1rem;
            margin-top: 10px;
        }}
        .metrics-bar {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }}
        .metric-box {{
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            color: white;
            padding: 25px;
            border-radius: 15px;
            text-align: center;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
            transition: transform 0.3s ease;
        }}
        .metric-box:hover {{
            transform: translateY(-5px);
        }}
        .metric-label {{
            font-size: 0.9rem;
            opacity: 0.9;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin: 0 0 10px 0;
        }}
        .metric-value {{
            font-size: 2rem;
            font-weight: bold;
            margin: 0;
        }}
        .key-insights {{
            background: linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%);
            padding: 25px;
            border-radius: 15px;
            margin: 20px 0;
            color: #333;
        }}
        .key-insights h2 {{
            margin: 0 0 15px 0;
            color: #d63384;
        }}
        .insight-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
        }}
        .insight-card {{
            background: rgba(255,255,255,0.9);
            padding: 20px;
            border-radius: 10px;
            text-align: center;
        }}
        .insight-number {{
            font-size: 2.5rem;
            font-weight: bold;
            color: #d63384;
            margin: 10px 0;
        }}
        .quick-wins {{
            background: #d4edda;
            border: 2px solid #28a745;
            border-radius: 15px;
            padding: 25px;
            margin: 20px 0;
        }}
        .quick-wins h2 {{
            color: #155724;
            margin: 0 0 15px 0;
        }}
        .hotel-card {{
            background: white;
            padding: 15px;
            margin: 10px 0;
            border-radius: 8px;
            border-left: 4px solid #28a745;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .hotel-info h4 {{
            margin: 0;
            color: #333;
        }}
        .hotel-info p {{
            margin: 5px 0 0 0;
            color: #666;
            font-size: 0.9rem;
        }}
        .issue-count {{
            background: #dc3545;
            color: white;
            padding: 8px 15px;
            border-radius: 20px;
            font-weight: bold;
        }}
        .pattern-insights {{
            background: #e3f2fd;
            border: 2px solid #2196f3;
            border-radius: 15px;
            padding: 25px;
            margin: 20px 0;
        }}
        .pattern-insights h2 {{
            color: #1976d2;
            margin: 0 0 15px 0;
        }}
        .pattern-card {{
            background: white;
            padding: 20px;
            margin: 10px 0;
            border-radius: 10px;
            border-left: 4px solid #2196f3;
        }}
        .pattern-card h4 {{
            margin: 0 0 10px 0;
            color: #1976d2;
        }}
        .impact-badge {{
            background: #2196f3;
            color: white;
            padding: 5px 10px;
            border-radius: 15px;
            font-size: 0.8rem;
            display: inline-block;
            margin-bottom: 10px;
        }}
        .action-plan {{
            background: #fff3cd;
            border: 2px solid #ffc107;
            border-radius: 15px;
            padding: 25px;
            margin: 20px 0;
        }}
        .action-plan h2 {{
            color: #856404;
            margin: 0 0 15px 0;
        }}
        .action-step {{
            background: white;
            padding: 15px;
            margin: 10px 0;
            border-radius: 8px;
            border-left: 4px solid #ffc107;
        }}
        .action-step h4 {{
            margin: 0 0 5px 0;
            color: #856404;
        }}
        .priority-high {{ color: #dc3545; font-weight: bold; }}
        .priority-medium {{ color: #ffc107; font-weight: bold; }}
        .priority-low {{ color: #28a745; font-weight: bold; }}
        .efficiency-meter {{
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin: 20px 0;
        }}
        .efficiency-bar {{
            background: #e9ecef;
            height: 20px;
            border-radius: 10px;
            overflow: hidden;
            margin: 10px 0;
        }}
        .efficiency-fill {{
            background: linear-gradient(90deg, #28a745, #20c997);
            height: 100%;
            transition: width 0.3s ease;
        }}
    </style>
</head>
<body>
<div class="dashboard">
    <div class="header">
        <h1>🎯 CRPM Action Dashboard</h1>
        <div class="subtitle">Focus on what matters most - {datetime.now().strftime('%B %d, %Y')}</div>
    </div>

    <!-- METRICS SUMMARY BAR -->
    <div class="metrics-bar">
        <div class="metric-box">
            <h3 class="metric-label">Total Rate Plans</h3>
            <p class="metric-value">{total_entries:,}</p>
        </div>
        <div class="metric-box">
            <h3 class="metric-label">Compliance Rate</h3>
            <p class="metric-value">{compliance_rate:.1f}%</p>
        </div>
        <div class="metric-box">
            <h3 class="metric-label">Issues to Fix</h3>
            <p class="metric-value">{issues_count:,}</p>
        </div>
        <div class="metric-box">
            <h3 class="metric-label">Accepted Deviations</h3>
            <p class="metric-value">{accepted_count:,}</p>
        </div>
    </div>

    <!-- KEY INSIGHTS SECTION -->
    <div class="key-insights">
        <h2>🚀 Key Insights - What You Need to Know</h2>
        <div class="insight-grid">
            <div class="insight-card">
                <div class="insight-number">{impact_percentage:.0f}%</div>
                <p><strong>Quick Win Impact</strong><br>Fixing top 5 hotels = {impact_percentage:.0f}% improvement</p>
            </div>
            <div class="insight-card">
                <div class="insight-number">{len(top_patterns)}</div>
                <p><strong>Automation Opportunities</strong><br>{sum(p[1]['count'] for p in top_patterns)} issues could be auto-fixed</p>
            </div>
            <div class="insight-card">
                <div class="insight-number">{rule_efficiency['efficiency_rate']:.0f}%</div>
                <p><strong>Current Automation</strong><br>{accepted_count:,} deviations auto-approved</p>
            </div>
            <div class="insight-card">
                <div class="insight-number">{len(top_problem_hotels)}</div>
                <p><strong>Focus Hotels</strong><br>These hotels need immediate attention
                </div>
            </div>
        </div>
        
        <!-- QUICK WINS SECTION -->
        <div class="quick-wins">
            <h2>🎯 Quick Wins - Fix These First (80/20 Rule)</h2>
            <p><strong>Strategy:</strong> These {len(top_problem_hotels[:5])} hotels account for {top_5_issues} issues ({impact_percentage:.0f}% of all problems). Fix these for maximum impact!</p>
            
            {f'''
            {"".join([f"""
            <div class="hotel-card">
                <div class="hotel-info">
                    <h4>Hotel {hotel}</h4>
                    <p>{len(hotel_details[hotel]['countries'])} countries • Priority: {max(hotel_details[hotel]['priorities'].items(), key=lambda x: x[1])[0] if hotel_details[hotel]['priorities'] else 'Medium'}</p>
                    <p><strong>Sample issues:</strong> {hotel_details[hotel]['issues'][0][:60]}{'...' if len(hotel_details[hotel]['issues'][0]) > 60 else ''}</p>
                </div>
                <div class="issue-count">{count} issues</div>
            </div>
            """ for hotel, count in top_problem_hotels[:5]])}
            ''' if top_problem_hotels else '''
            <div class="insight-card">
                <h3>🎉 Excellent Distribution!</h3>
                <p>No single hotel dominates your issues. Problems are well-distributed, indicating good overall management.</p>
            </div>
            '''}
        </div>
        
        <!-- AUTOMATION OPPORTUNITIES -->
        <div class="pattern-insights">
            <h2>🤖 Automation Opportunities - Create These Rules</h2>
            <p><strong>Smart Suggestion:</strong> These patterns appear frequently. Create exception rules to handle them automatically.</p>
            
            {f'''
            {"".join([f"""
            <div class="pattern-card">
                <div class="impact-badge">Would fix {data['count']} issues automatically</div>
                <h4>{pattern.replace('_', ' ').title()}</h4>
                <p><strong>Scope:</strong> {len(data['hotels'])} hotels across {len(data['countries'])} countries</p>
                <p><strong>Examples:</strong> {', '.join(data['examples'])}</p>
                <p><strong>Recommended Rule:</strong> Create a {'Country_Pattern' if len(data['countries']) == 1 else 'Hotel_Pattern'} exception for this deviation</p>
            </div>
            """ for pattern, data in top_patterns])}
            ''' if top_patterns else '''
            <div class="insight-card">
                <h3>🔧 Well Optimized!</h3>
                <p>No major automation opportunities found. Your exception rules are handling patterns effectively.</p>
            </div>
            '''}
        </div>
        
        <!-- REGIONAL FOCUS -->
        <div class="pattern-insights">
            <h2>🌍 Regional Focus - Where to Concentrate Efforts</h2>
            {f'''
            {"".join([f"""
            <div class="pattern-card">
                <h4>{country} - {data['issues']} issues</h4>
                <p><strong>Hotels affected:</strong> {len(data['hotels'])} hotels need attention</p>
                <p><strong>Main problems:</strong> {', '.join([f"{issue_type} ({count})" for issue_type, count in sorted(data['types'].items(), key=lambda x: x[1], reverse=True)])}</p>
                <p><strong>Action:</strong> {'Focus on VAT standardization' if data['types'].get('VAT', 0) > data['issues']/2 else 'Review subaccount configurations' if data['types'].get('Subaccount', 0) > data['issues']/2 else 'Mixed issues - needs detailed review'}</p>
            </div>
            """ for country, data in top_countries])}
            ''' if top_countries else '''
            <div class="insight-card">
                <h3>🌟 Global Excellence!</h3>
                <p>Issues are well-distributed across regions. No single country requires special focus.</p>
            </div>
            '''}
        </div>
        
        <!-- ACTION PLAN -->
        <div class="action-plan">
            <h2>📋 Your 30-Day Action Plan</h2>
            
            <div class="action-step">
                <h4>Week 1: Quick Wins (Immediate Impact)</h4>
                <p>• Review top {min(3, len(top_problem_hotels))} hotels: {', '.join([hotel for hotel, _ in top_problem_hotels[:3]])}</p>
                <p>• Expected impact: Fix ~{sum(count for _, count in top_problem_hotels[:3])} issues ({sum(count for _, count in top_problem_hotels[:3])/issues_count*100:.0f}% improvement)</p>
            </div>
            
            <div class="action-step">
                <h4>Week 2-3: Automation Rules</h4>
                <p>• Implement {min(2, len(top_patterns))} new exception rules for top patterns</p>
                <p>• Expected impact: Auto-handle {sum(data['count'] for _, data in top_patterns[:2])} future similar issues</p>
            </div>
            
            <div class="action-step">
                <h4>Week 4: Regional Focus</h4>
                <p>• Address {top_countries[0][0] if top_countries else 'regional'} issues systematically</p>
                <p>• Set up monitoring for early detection of similar patterns</p>
            </div>
            
            <div class="action-step">
                <h4>Ongoing: Monitor & Optimize</h4>
                <p>• Target: Achieve {min(98, compliance_rate + 5):.0f}% compliance rate</p>
                <p>• Review exception rule effectiveness monthly</p>
            </div>
        </div>
        
        <!-- EFFICIENCY TRACKING -->
        <div class="efficiency-meter">
            <h3>🚀 Automation Efficiency</h3>
            <p>Your exception rules are handling <strong>{rule_efficiency['efficiency_rate']:.0f}%</strong> of deviations automatically</p>
            <div class="efficiency-bar">
                <div class="efficiency-fill" style="width: {rule_efficiency['efficiency_rate']}%"></div>
            </div>
            <p><small>{accepted_count:,} auto-approved • {issues_count:,} need manual review • Target: 95% automation</small></p>
        </div>
        
        <div style="text-align: center; margin-top: 30px; color: #666; font-size: 0.9rem;">
            <p>📊 Actionable CRPM Dashboard - Focus on Impact, Not Just Data</p>
            <p>Next recommended review: {(datetime.now() + timedelta(days=7)).strftime('%B %d, %Y')}</p>
        </div>
    </div>
</body>
</html>
"""
    dashboard_file = os.path.join(output_dir, f"CRPM_Dashboard_{timestamp}.html")
    with open(dashboard_file, 'w', encoding='utf-8') as f:
        f.write(html_content)
    return dashboard_file

print("🚀 Starting SIMPLIFIED ENHANCED CRPM analysis...")
