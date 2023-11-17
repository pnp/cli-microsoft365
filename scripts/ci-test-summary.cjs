const mocha = require('mocha');
const core = require('@actions/core');

const { EVENT_RUN_END, EVENT_TEST_FAIL, EVENT_RUN_BEGIN } =
  mocha.Runner.constants;

class TestSummaryReporter {
  summary = core.summary;
  testResult = { failedTests: {}, stats: {} };
  suiteIndenter = '&emsp; ';
  destination;

  constructor(runner, options) {
    this._indents = 0;
    const stats = runner.stats;

    new mocha.reporters.Min(runner, options);

    runner
      .on(EVENT_RUN_BEGIN, () => {
        if (process.env.GITHUB_STEP_SUMMARY) {
          // Capturing the value of GITHUB_STEP_SUMMARY to use later
          this.destination = process.env.GITHUB_STEP_SUMMARY;
        } else {
          console.error(
            `Destination not found, please ensure GITHUB_STEP_SUMMARY is set and valid. GITHUB_STEP_SUMMARY: ${process.env.GITHUB_STEP_SUMMARY}`
          );
          process.exit(1);
        }
      })
      .on(EVENT_TEST_FAIL, (test, _) => {
        const testPath = test.titlePath();
        const testName = testPath.pop();
        let resultPath = this.testResult.failedTests;
        testPath.forEach((item, index) => {
          if (index === testPath.length - 1) {
            if (!resultPath.hasOwnProperty(item)) {
              resultPath[item] = [];
            }
            resultPath[item].push(testName);
          }
          if (!resultPath.hasOwnProperty(item)) {
            resultPath[item] = {};
          }
          resultPath = resultPath[item];
        });
      })
      .once(EVENT_RUN_END, () => {
        this.testResult.stats = {
          passed: stats.passes,
          failed: stats.failures,
          total: stats.passes + stats.failures,
        };
        this.writeSummary();
      });
  }

  async writeSummary() {
    try {
      if (this.testResult.stats.failed > 0) {
        this.summary = this.summary.addHeading('Failed tests', 3);
        this.summary = this.summary
          .addRaw(this.generateAccordian('', null, this.testResult.failedTests))
          .addSeparator();
      }
      this.writeStatsTable();
      this.writeProgressBar();
      // Resettig GITHUB_STEP_SUMMARY, since unit tests might overwrite it
      process.env.GITHUB_STEP_SUMMARY = this.destination;
      await this.summary.write();
    } catch (e) {
      console.error(e);
    }
  }

  generateAccordian(suitePrefix, suiteName, suiteObject) {
    if (typeof suiteObject !== 'object') {
      // Item is a test, not a suite
      return `${suitePrefix}❌ ${suiteObject}<br />`;
    }
    let detailsContent = '';

    // Generate accordian for each item in suite
    Object.keys(suiteObject).forEach(
      (item) =>
        (detailsContent =
          detailsContent +
          this.generateAccordian(
            `${suitePrefix}${this.suiteIndenter}`,
            item,
            suiteObject[item]
          ))
    );

    // First level does not require accordian
    if (suiteName === null) {
      return detailsContent;
    }

    // Return an accordian for the suite
    return `
      <details>
        <summary>${suitePrefix}${suiteName}</summary>
        ${detailsContent}
      </details>
    `;
  }

  writeStatsTable() {
    this.summary.addHeading('Stats', 3);
    this.summary = this.summary.addTable([
      ['Total number of tests', '' + this.testResult.stats.total],
      ['Passed ✅', '' + this.testResult.stats.passed],
      ['Failed ❌', '' + this.testResult.stats.failed],
    ]);
  }

  writeProgressBar() {
    let percentagePassed = Math.floor(
      (this.testResult.stats.passed / this.testResult.stats.total) * 100
    );
    this.summary = this.summary.addRaw(`
      <div style="width: 50%; background-color: red; text-align: center; color: white;">
        <div style="width: ${percentagePassed}%; height: 18px; background-color: green;"></div>
      </div>  
    `);
  }
}

module.exports = TestSummaryReporter;
