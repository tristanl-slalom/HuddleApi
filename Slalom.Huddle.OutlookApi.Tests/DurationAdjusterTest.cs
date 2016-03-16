using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Slalom.Huddle.OutlookApi.Services;
using System.Linq;

namespace Slalom.Huddle.OutlookApi.Tests
{
    [TestClass]
    public class DurationAdjusterTest
    {
        [TestMethod]
        public void TestCurrentTime()
        {
            // Arrange
            int duration = 30;
            DurationAdjuster durationAdjuster = new DurationAdjuster();

            // Act
            DateTime actual = durationAdjuster.ExtendDurationToNearestBlock(duration);

            // Assert
            Assert.IsTrue(new [] { 0, 15, 30, 45}.Contains(actual.Minute));
            Assert.IsTrue(actual.Second == 0);
            Assert.IsTrue(actual.Millisecond == 0);
        }
    }
}
