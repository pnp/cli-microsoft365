const mocha = require('mocha');
const core = import('@actions/core');

const { EVENT_RUN_END, EVENT_TEST_FAIL, EVENT_RUN_BEGIN } =
  mocha.Runner.constants;

class TestSummaryReporter {
  summary = core.summary;
  testResult = { failedTests: {}, stats: {} };
  suiteIndenter = '&emsp; ';
  destination;

  constructor(runner, options) {
    new mocha.reporters.Min(runner, options);

    if (!process.env.GITHUB_STEP_SUMMARY) {
      return;
    }

    const stats = runner.stats;
    runner
      .on(EVENT_RUN_BEGIN, () => {
        // Capturing the value of GITHUB_STEP_SUMMARY to use later
        this.destination = process.env.GITHUB_STEP_SUMMARY;
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
          .addRaw(this.generateAccordion('', null, this.testResult.failedTests))
          .addSeparator();
      }
      this.writeStatsTable();
      this.writeResultsBar();
      // Resettig GITHUB_STEP_SUMMARY, since unit tests might overwrite it
      process.env.GITHUB_STEP_SUMMARY = this.destination;
      await this.summary.write();
    } catch (e) {
      console.error(e);
    }
  }

  generateAccordion(suitePrefix, suiteName, suiteObject) {
    if (typeof suiteObject !== 'object') {
      // Item is a test, not a suite
      return `${suitePrefix}❌ ${suiteObject}<br />`;
    }
    let detailsContent = '';

    // Generate accordion for each item in suite
    Object.keys(suiteObject).forEach(
      (item) =>
      (detailsContent =
        detailsContent +
        this.generateAccordion(
          `${suitePrefix}${this.suiteIndenter}`,
          item,
          suiteObject[item]
        ))
    );

    // First level does not require accordion
    if (suiteName === null) {
      return detailsContent;
    }

    // Return an accordion for the suite
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

  writeResultsBar() {
    let percentagePassed = Math.floor(
      (this.testResult.stats.passed / this.testResult.stats.total) * 100
    );
    this.summary = this.summary.addRaw(
      '<div style="width: 50%; background-color: red; text-align: center; color: white;"><div style="width:' +
      percentagePassed +
      '%; height: 18px; background-color: green;"></div></div>'
    );
  }
}

module.exports = TestSummaryReporter;
