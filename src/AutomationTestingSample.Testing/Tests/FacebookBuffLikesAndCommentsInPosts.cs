using AutomationTestingSample.Testing.Pages;

namespace AutomationTestingSample.Testing.Tests
{
    public class FacebookBuffLikesAndCommentsInPosts : TestBase
    {
        public FacebookBuffLikesAndCommentsInPosts(string browserType, int browserWidth, int browserHeight) : base(browserType, browserWidth, browserHeight)
        {
        }

        [TestCase]
        public void AutoBuffLikesAndComments()
        {
            var page = new BuffLikesAndCommentsPage(Driver);
            page.BuffLikes();
        }
    }
}
